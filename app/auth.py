"""
Sistema de autenticación y autorización
"""
from flask import Blueprint, render_template, request, redirect, url_for, flash, Response
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash
from app.models import db, User, ReportHistory
from config import USER_ROLES, PERMISSIONS
import logging

logger = logging.getLogger(__name__)

auth_bp = Blueprint('auth', __name__)

def init_auth(app):
    """Inicializa el sistema de autenticación"""
    from flask_login import LoginManager
    
    login_manager = LoginManager()
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'
    login_manager.login_message = 'Por favor inicia sesión para acceder a esta página.'
    login_manager.login_message_category = 'info'
    
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))
    
    # Inicializar base de datos
    with app.app_context():
        db.create_all()
        create_default_users()

def create_default_users():
    """Crea usuarios por defecto si no existen"""
    if User.query.count() == 0:
        # Usuario administrador por defecto
        admin = User(
            username='admin',
            email='admin@crediflexi.com',
            role='ADMIN'
        )
        admin.set_password('admin123')
        db.session.add(admin)
        
        # Usuario consultor por defecto
        consultor = User(
            username='consultor',
            email='consultor@crediflexi.com',
            role='CONSULTOR'
        )
        consultor.set_password('consultor123')
        db.session.add(consultor)
        
        db.session.commit()
        logger.info("Usuarios por defecto creados: admin/admin123, consultor/consultor123")

def require_permission(permission):
    """Decorador para requerir permisos específicos"""
    def decorator(f):
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for('auth.login'))
            if not current_user.has_permission(permission):
                flash('No tienes permisos para acceder a esta función.', 'error')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        decorated_function.__name__ = f.__name__
        return decorated_function
    return decorator

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    """Página de inicio de sesión"""
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if not username or not password:
            flash('Por favor completa todos los campos.', 'error')
            return render_template('auth/login.html')
        
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password) and user.is_active:
            login_user(user)
            user.last_login = db.func.now()
            db.session.commit()
            
            flash(f'¡Bienvenido, {user.username}!', 'success')
            
            # Redirigir según el rol
            if user.is_admin():
                return redirect(url_for('index'))
            else:
                return redirect(url_for('consultor.dashboard'))
        else:
            flash('Usuario o contraseña incorrectos.', 'error')
    
    return render_template('auth/login.html')

@auth_bp.route('/logout')
@login_required
def logout():
    """Cerrar sesión"""
    username = current_user.username
    logout_user()
    flash(f'¡Hasta luego, {username}!', 'info')
    return redirect(url_for('auth.login'))

@auth_bp.route('/register', methods=['GET', 'POST'])
@login_required
@require_permission('manage_users')
def register():
    """Registro de nuevos usuarios (solo administradores)"""
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        role = request.form.get('role')
        
        if not all([username, email, password, role]):
            flash('Por favor completa todos los campos.', 'error')
            return render_template('auth/register.html', roles=USER_ROLES)
        
        # Verificar si el usuario ya existe
        if User.query.filter_by(username=username).first():
            flash('El nombre de usuario ya existe.', 'error')
            return render_template('auth/register.html', roles=USER_ROLES)
        
        if User.query.filter_by(email=email).first():
            flash('El email ya está registrado.', 'error')
            return render_template('auth/register.html', roles=USER_ROLES)
        
        # Crear nuevo usuario
        user = User(
            username=username,
            email=email,
            role=role
        )
        user.set_password(password)
        
        try:
            db.session.add(user)
            db.session.commit()
            flash(f'Usuario {username} creado exitosamente.', 'success')
            return redirect(url_for('auth.users'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error al crear usuario: {str(e)}', 'error')
    
    return render_template('auth/register.html', roles=USER_ROLES)

@auth_bp.route('/users')
@login_required
@require_permission('manage_users')
def users():
    """Lista de usuarios (solo administradores)"""
    users = User.query.all()
    return render_template('auth/users.html', users=users, roles=USER_ROLES)

@auth_bp.route('/profile')
@login_required
def profile():
    """Perfil del usuario actual"""
    return render_template('auth/profile.html', user=current_user, roles=USER_ROLES)
