from flask import Flask, render_template
from app.reportes import reportes_bp

def create_app():
    """Factory function para crear la aplicación Flask"""
    app = Flask(__name__)
    
    # Configuración básica
    app.config['SECRET_KEY'] = 'tu-clave-secreta-aqui'
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
    app.config['UPLOAD_FOLDER'] = 'uploads'
    
    # Registrar blueprints
    app.register_blueprint(reportes_bp, url_prefix='/reportes')
    
    @app.route('/')
    def index():
        """Página principal con opciones de reportes"""
        return render_template('index.html')
    
    return app

if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, host='0.0.0.0', port=5000)
