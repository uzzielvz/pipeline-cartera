from flask import Flask, render_template, redirect, url_for
from flask_login import login_required, current_user
from app.reportes import reportes_bp
from app.auth import auth_bp, init_auth
from app.consultor import consultor_bp
from app.models import db
from config import SECRET_KEY, SQLALCHEMY_DATABASE_URI, SQLALCHEMY_TRACK_MODIFICATIONS, UPLOAD_FOLDER, MAX_FILE_SIZE

def create_app():
    """Factory function para crear la aplicación Flask"""
    app = Flask(__name__)
    
    # Configuración básica
    app.config['SECRET_KEY'] = SECRET_KEY
    app.config['SQLALCHEMY_DATABASE_URI'] = SQLALCHEMY_DATABASE_URI
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = SQLALCHEMY_TRACK_MODIFICATIONS
    app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
    
    # Inicializar extensiones
    db.init_app(app)
    init_auth(app)
    
    # Registrar blueprints
    app.register_blueprint(auth_bp, url_prefix='/auth')
    app.register_blueprint(reportes_bp, url_prefix='/reportes')
    app.register_blueprint(consultor_bp, url_prefix='/consultor')
    
    @app.route('/')
    @login_required
    def index():
        """Página principal con opciones de reportes"""
        if current_user.is_admin():
            return render_template('index.html')
        else:
            return redirect(url_for('consultor.dashboard'))
    
    @app.route('/dashboard')
    @login_required
    def dashboard():
        """Dashboard accesible para admin y consultor"""
        return redirect(url_for('consultor.dashboard'))
    
    @app.route('/unauthorized')
    def unauthorized():
        """Página de acceso no autorizado"""
        return render_template('errors/unauthorized.html'), 403
    
    return app

if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, host='0.0.0.0', port=5000)
