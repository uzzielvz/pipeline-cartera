"""
Modelos de base de datos para el sistema de autenticación
"""
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

db = SQLAlchemy()

class User(UserMixin, db.Model):
    """Modelo de usuario con autenticación"""
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='CONSULTOR')
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)
    
    def __repr__(self):
        return f'<User {self.username}>'
    
    def set_password(self, password):
        """Establece la contraseña hasheada"""
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        """Verifica la contraseña"""
        return check_password_hash(self.password_hash, password)
    
    def has_permission(self, permission):
        """Verifica si el usuario tiene un permiso específico"""
        from config import PERMISSIONS
        return permission in PERMISSIONS.get(self.role, [])
    
    def is_admin(self):
        """Verifica si el usuario es administrador"""
        return self.role == 'ADMIN'
    
    def is_consultor(self):
        """Verifica si el usuario es consultor"""
        return self.role == 'CONSULTOR'

class ReportHistory(db.Model):
    """Historial de reportes generados"""
    __tablename__ = 'report_history'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    report_type = db.Column(db.String(50), nullable=False)  # 'individual' o 'grupal'
    filename = db.Column(db.String(255), nullable=False)
    file_path = db.Column(db.String(500), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    file_size = db.Column(db.Integer)
    
    # Relación con usuario
    user = db.relationship('User', backref=db.backref('reports', lazy=True))
    
    def __repr__(self):
        return f'<ReportHistory {self.filename}>'