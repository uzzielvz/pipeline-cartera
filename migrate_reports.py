#!/usr/bin/env python3
"""
Script para migrar reportes existentes al nuevo directorio de reportes
"""
import sys
import os
import shutil
from datetime import datetime

# Agregar el directorio del proyecto al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask
from app.models import db, ReportHistory
from config import REPORTS_FOLDER, SECRET_KEY, SQLALCHEMY_DATABASE_URI, SQLALCHEMY_TRACK_MODIFICATIONS

def migrate_existing_reports():
    """Migrar reportes existentes al nuevo directorio"""
    
    # Crear una instancia de Flask temporal
    app = Flask(__name__)
    app.config['SECRET_KEY'] = SECRET_KEY
    app.config['SQLALCHEMY_DATABASE_URI'] = SQLALCHEMY_DATABASE_URI
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = SQLALCHEMY_TRACK_MODIFICATIONS
    
    # Inicializar la base de datos
    db.init_app(app)
    
    with app.app_context():
        # Crear directorio de reportes si no existe
        os.makedirs(REPORTS_FOLDER, exist_ok=True)
        
        # Obtener todos los reportes de la base de datos
        reports = ReportHistory.query.all()
        
        migrated_count = 0
        error_count = 0
        
        print(f"🔍 Encontrados {len(reports)} reportes para migrar...")
        
        for report in reports:
            try:
                old_path = report.file_path
                
                # Verificar si el archivo existe
                if not os.path.exists(old_path):
                    print(f"⚠️  Archivo no encontrado: {old_path}")
                    error_count += 1
                    continue
                
                # Generar nueva ruta
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = os.path.basename(old_path)
                name, ext = os.path.splitext(filename)
                
                # Crear nuevo nombre con timestamp y tipo
                new_filename = f"{name}_{report.report_type}_migrated_{timestamp}{ext}"
                new_path = os.path.join(REPORTS_FOLDER, new_filename)
                
                # Mover archivo
                shutil.copy2(old_path, new_path)
                
                # Actualizar base de datos
                report.file_path = new_path
                report.filename = new_filename
                
                migrated_count += 1
                print(f"✅ Migrado: {filename} -> {new_filename}")
                
            except Exception as e:
                print(f"❌ Error migrando {report.filename}: {str(e)}")
                error_count += 1
        
        # Guardar cambios en la base de datos
        try:
            db.session.commit()
            print(f"\n🎉 Migración completada:")
            print(f"   ✅ Reportes migrados: {migrated_count}")
            print(f"   ❌ Errores: {error_count}")
            print(f"   📁 Directorio: {REPORTS_FOLDER}")
        except Exception as e:
            db.session.rollback()
            print(f"❌ Error guardando cambios en la base de datos: {str(e)}")

if __name__ == '__main__':
    print("🚀 Iniciando migración de reportes...")
    migrate_existing_reports()
    print("✅ Migración finalizada.")
