"""
Blueprint para usuarios consultores - solo pueden ver reportes
"""
from flask import Blueprint, render_template, send_file, request, flash, redirect, url_for
from flask_login import login_required, current_user
from app.models import db, ReportHistory, User
from config import USER_ROLES
import os
import logging

logger = logging.getLogger(__name__)

consultor_bp = Blueprint('consultor', __name__)

@consultor_bp.route('/dashboard')
@login_required
def dashboard():
    """Dashboard accesible para consultores y administradores"""
    
    # Obtener reportes recientes (últimos 30 días)
    from datetime import datetime, timedelta
    thirty_days_ago = datetime.utcnow() - timedelta(days=30)
    
    recent_reports = ReportHistory.query.filter(
        ReportHistory.created_at >= thirty_days_ago
    ).order_by(ReportHistory.created_at.desc()).limit(20).all()
    
    # Estadísticas básicas
    total_reports = ReportHistory.query.count()
    individual_reports = ReportHistory.query.filter_by(report_type='individual').count()
    grupal_reports = ReportHistory.query.filter_by(report_type='grupal').count()
    
    stats = {
        'total_reports': total_reports,
        'individual_reports': individual_reports,
        'grupal_reports': grupal_reports,
        'recent_reports': len(recent_reports)
    }
    
    return render_template('consultor/dashboard.html', 
                         reports=recent_reports, 
                         stats=stats)

@consultor_bp.route('/reports')
@login_required
def reports():
    """Lista todos los reportes disponibles"""
    
    page = request.args.get('page', 1, type=int)
    per_page = 20
    
    reports = ReportHistory.query.order_by(
        ReportHistory.created_at.desc()
    ).paginate(
        page=page, 
        per_page=per_page, 
        error_out=False
    )
    
    return render_template('consultor/reports.html', reports=reports)

@consultor_bp.route('/download/<int:report_id>')
@login_required
def download_report(report_id):
    """Descargar un reporte específico"""
    
    report = ReportHistory.query.get_or_404(report_id)
    
    if not os.path.exists(report.file_path):
        flash('El archivo del reporte no existe.', 'error')
        return redirect(url_for('consultor.reports'))
    
    try:
        return send_file(
            report.file_path,
            as_attachment=True,
            download_name=report.filename
        )
    except Exception as e:
        logger.error(f"Error descargando reporte {report_id}: {str(e)}")
        flash('Error al descargar el archivo.', 'error')
        return redirect(url_for('consultor.reports'))

@consultor_bp.route('/report/<int:report_id>')
@login_required
def view_report(report_id):
    """Ver detalles de un reporte específico"""
    
    report = ReportHistory.query.get_or_404(report_id)
    user = User.query.get(report.user_id)
    
    return render_template('consultor/report_detail.html', 
                         report=report, 
                         user=user)

@consultor_bp.route('/delete_report/<int:report_id>', methods=['POST'])
@login_required
def delete_report(report_id):
    """Eliminar un reporte (solo administradores)"""
    if not current_user.is_admin():
        flash('No tienes permisos para eliminar reportes.', 'error')
        return redirect(url_for('consultor.dashboard'))
    
    report = ReportHistory.query.get_or_404(report_id)
    
    try:
        # Eliminar archivo físico si existe
        if os.path.exists(report.file_path):
            os.remove(report.file_path)
        
        # Eliminar registro de la base de datos
        db.session.delete(report)
        db.session.commit()
        
        flash(f'Reporte "{report.filename}" eliminado exitosamente.', 'success')
        logger.info(f"Reporte {report_id} eliminado por usuario {current_user.username}")
        
    except Exception as e:
        db.session.rollback()
        flash(f'Error al eliminar el reporte: {str(e)}', 'error')
        logger.error(f"Error eliminando reporte {report_id}: {str(e)}")
    
    return redirect(url_for('consultor.dashboard'))
