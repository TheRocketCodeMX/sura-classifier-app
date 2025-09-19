#!/usr/bin/env python3
"""
Email Classification Web Dashboard
Aplicación web para visualizar y analizar emails clasificados
"""

from flask import Flask, render_template, request, jsonify, send_file, abort
import json
import pandas as pd
import plotly.graph_objs as go
import plotly.utils
from pathlib import Path
import email
import zipfile
import io
from datetime import datetime
import re
import mimetypes
import os

# Configuración de Google Drive
try:
    from config_drive import GOOGLE_DRIVE_FOLDER_URL
except ImportError:
    GOOGLE_DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/TU_FOLDER_ID_AQUI"

app = Flask(__name__)

class EmailDashboard:
    def __init__(self, output_dir="output"):
        self.output_dir = Path(output_dir)
        self.classification_file = self.output_dir / "classification" / "classification_results.json"
        self.emails_dir = self.output_dir / "emails"
        self.attachments_dir = self.output_dir / "attachments"
        self.metadata_dir = self.output_dir / "metadata"

        # Cargar datos de clasificación
        self.load_classification_data()

    def load_classification_data(self):
        """Cargar datos de clasificación"""
        try:
            if not self.classification_file.exists():
                print(f"Archivo de clasificación no encontrado: {self.classification_file}")
                self.data = {'emails': []}
                self.df = pd.DataFrame()
                return

            with open(self.classification_file, 'r', encoding='utf-8') as f:
                self.data = json.load(f)

            # Convertir a DataFrame para análisis
            emails_data = []
            for email_info in self.data.get('emails', []):
                row = {
                    'email_id': email_info['email_id'],
                    'subject': email_info['metadata'].get('subject', ''),
                    'sender_name': email_info['metadata'].get('sender_name', ''),
                    'sender_email': email_info['metadata'].get('sender_email', ''),
                    'folder': email_info['metadata'].get('folder', ''),
                    'delivery_time': email_info['metadata'].get('delivery_time', ''),
                    'size': email_info['metadata'].get('size', 0),
                    'attachment_count': email_info['metadata'].get('attachment_count', 0),
                    'classification_type': email_info['primary_classification']['type'],
                    'confidence': email_info['primary_classification']['confidence'],
                    'status': email_info['primary_classification']['status'],
                    'agente_code': email_info['primary_classification']['details'].get('agente_code', ''),
                    'poliza_number': email_info['primary_classification']['details'].get('poliza_number', ''),
                    'has_slip': email_info['attachment_analysis'].get('has_slip', False),
                    'slip_complete': email_info['attachment_analysis'].get('slip_complete', False),
                    'total_attachments': email_info['attachment_analysis'].get('total_attachments', 0)
                }
                emails_data.append(row)

            self.df = pd.DataFrame(emails_data)

            # Limpiar fechas
            self.df['delivery_date'] = pd.to_datetime(self.df['delivery_time'], errors='coerce')

            # Convertir fechas NaT a None para evitar errores de serialización JSON
            self.df['delivery_time'] = self.df['delivery_time'].where(pd.notnull(self.df['delivery_time']), None)

        except Exception as e:
            print(f"Error cargando datos: {e}")
            self.data = {'emails': []}
            self.df = pd.DataFrame()

    def get_summary_stats(self):
        """Obtener estadísticas resumen"""
        total = len(self.df)

        if total == 0:
            return {}

        stats = {
            'total_emails': total,
            'cotizacion': len(self.df[self.df['classification_type'] == 'cotizacion']),
            'renovacion': len(self.df[self.df['classification_type'] == 'renovacion']),
            'endoso': len(self.df[self.df['classification_type'] == 'endoso']),
            'sin_clasificar': len(self.df[self.df['classification_type'] == 'sin_clasificar']),
            'with_attachments': len(self.df[self.df['total_attachments'] > 0]),
            'with_slip': len(self.df[self.df['has_slip'] == True]),
            'complete_slip': len(self.df[self.df['slip_complete'] == True])
        }

        return stats

    def create_charts(self):
        """Crear gráficos para el dashboard"""
        charts = {}

        if len(self.df) == 0:
            return charts

        # Gráfico de clasificación
        class_counts = self.df['classification_type'].value_counts()

        charts['classification_pie'] = {
            'data': [go.Pie(
                labels=class_counts.index,
                values=class_counts.values,
                hole=0.3
            )],
            'layout': go.Layout(
                title='Distribución por Clasificación',
                height=400
            )
        }

        # Gráfico de emails por fecha
        if 'delivery_date' in self.df.columns:
            df_with_dates = self.df.dropna(subset=['delivery_date'])
            if len(df_with_dates) > 0:
                df_with_dates = df_with_dates.copy()
                df_with_dates['date_only'] = df_with_dates['delivery_date'].dt.date
                date_counts = df_with_dates.groupby(['date_only', 'classification_type']).size().reset_index(name='count')

                charts['timeline'] = {
                    'data': [],
                    'layout': go.Layout(
                        title='Emails por Fecha y Clasificación',
                        height=400,
                        xaxis={'title': 'Fecha'},
                        yaxis={'title': 'Cantidad de Emails'}
                    )
                }

                for class_type in date_counts['classification_type'].unique():
                    class_data = date_counts[date_counts['classification_type'] == class_type]
                    charts['timeline']['data'].append(
                        go.Scatter(
                            x=class_data['date_only'],
                            y=class_data['count'],
                            mode='lines+markers',
                            name=class_type
                        )
                    )

        # Gráfico de confianza por clasificación
        conf_data = []
        for class_type in ['cotizacion', 'renovacion', 'endoso']:
            class_df = self.df[self.df['classification_type'] == class_type]
            if len(class_df) > 0:
                conf_data.append(go.Box(
                    y=class_df['confidence'],
                    name=class_type,
                    boxpoints='outliers'
                ))

        if conf_data:
            charts['confidence_box'] = {
                'data': conf_data,
                'layout': go.Layout(
                    title='Distribución de Confianza por Clasificación',
                    height=400,
                    yaxis={'title': 'Confianza (%)'}
                )
            }

        return charts

    def search_emails(self, query="", classification="", folder="", has_attachments=None, date_from="", date_to="", page=1, per_page=50):
        """Buscar emails con filtros y paginación"""
        df_filtered = self.df.copy()

        # Filtrar por query en asunto
        if query:
            df_filtered = df_filtered[
                df_filtered['subject'].str.contains(query, case=False, na=False) |
                df_filtered['sender_name'].str.contains(query, case=False, na=False)
            ]

        # Filtrar por clasificación
        if classification and classification != 'all':
            df_filtered = df_filtered[df_filtered['classification_type'] == classification]

        # Filtrar por carpeta
        if folder and folder != 'all':
            df_filtered = df_filtered[df_filtered['folder'] == folder]

        # Filtrar por adjuntos
        if has_attachments is not None:
            if has_attachments:
                df_filtered = df_filtered[df_filtered['total_attachments'] > 0]
            else:
                df_filtered = df_filtered[df_filtered['total_attachments'] == 0]

        # Filtrar por fechas
        if date_from:
            df_filtered = df_filtered[df_filtered['delivery_date'] >= date_from]
        if date_to:
            df_filtered = df_filtered[df_filtered['delivery_date'] <= date_to]

        # Calcular paginación
        total_results = len(df_filtered)
        total_pages = (total_results + per_page - 1) // per_page
        start_idx = (page - 1) * per_page
        end_idx = start_idx + per_page

        # Aplicar paginación
        df_paginated = df_filtered.iloc[start_idx:end_idx]

        # Convertir a diccionario y limpiar valores problemáticos
        results = df_paginated.to_dict('records')

        # Limpiar fechas NaT en los resultados
        for result in results:
            if pd.isna(result.get('delivery_time')):
                result['delivery_time'] = None
            if pd.isna(result.get('delivery_date')):
                result['delivery_date'] = None

        return {
            'results': results,
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total': total_results,
                'pages': total_pages,
                'has_prev': page > 1,
                'has_next': page < total_pages,
                'prev_page': page - 1 if page > 1 else None,
                'next_page': page + 1 if page < total_pages else None
            }
        }

    def get_file_type_info(self, filename):
        """Obtener información del tipo de archivo"""
        # Detectar tipo MIME
        mime_type, _ = mimetypes.guess_type(filename)

        # Obtener extensión
        _, ext = os.path.splitext(filename.lower())

        # Definir categorías (MINIMALISTA)
        file_info = {
            'category': 'Documento',
            'mime_type': mime_type or 'application/octet-stream',
            'extension': ext,
            'color_class': 'bg-secondary'
        }

        # Mapeo de tipos de archivo (MINIMALISTA - SIN ICONOS)
        if ext in ['.pdf']:
            file_info.update({
                'category': 'PDF',
                'color_class': 'bg-danger'
            })
        elif ext in ['.doc', '.docx']:
            file_info.update({
                'category': 'Word',
                'color_class': 'bg-primary'
            })
        elif ext in ['.xls', '.xlsx']:
            file_info.update({
                'category': 'Excel',
                'color_class': 'bg-success'
            })
        elif ext in ['.ppt', '.pptx']:
            file_info.update({
                'category': 'PowerPoint',
                'color_class': 'bg-warning'
            })
        elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']:
            file_info.update({
                'category': 'Imagen',
                'color_class': 'bg-info'
            })
        elif ext in ['.zip', '.rar', '.7z']:
            file_info.update({
                'category': 'Archivo',
                'color_class': 'bg-dark text-white'
            })
        elif ext in ['.txt', '.log']:
            file_info.update({
                'category': 'Texto',
                'color_class': 'bg-light text-dark'
            })
        elif ext in ['.csv']:
            file_info.update({
                'category': 'CSV',
                'color_class': 'bg-success'
            })
        elif ext in ['.xml', '.html', '.htm']:
            file_info.update({
                'category': 'Web',
                'color_class': 'bg-info'
            })

        # Detectar archivos especiales de seguros (SIN ICONOS - MINIMALISTA)
        filename_upper = filename.upper()
        if 'SLIP' in filename_upper:
            file_info.update({
                'category': 'SLIP',
                'color_class': 'bg-primary'
            })
        elif any(word in filename_upper for word in ['POLIZA', 'POLICY']):
            file_info.update({
                'category': 'Póliza',
                'color_class': 'bg-warning'
            })
        elif any(word in filename_upper for word in ['COTIZACION', 'QUOTE']):
            file_info.update({
                'category': 'Cotización',
                'color_class': 'bg-success'
            })
        elif any(word in filename_upper for word in ['ENDOSO', 'ENDORSEMENT']):
            file_info.update({
                'category': 'Endoso',
                'color_class': 'bg-info'
            })

        return file_info

    def get_email_content(self, email_id):
        """Obtener contenido completo del email"""
        try:
            # Cargar metadatos
            metadata_file = self.metadata_dir / f"{email_id}.json"
            with open(metadata_file, 'r', encoding='utf-8') as f:
                metadata = json.load(f)

            # Cargar contenido del .eml
            eml_file = self.emails_dir / f"{email_id}.eml"
            content = {"plain_text": "", "html_content": ""}

            if eml_file.exists():
                with open(eml_file, 'r', encoding='utf-8') as f:
                    msg = email.message_from_file(f)

                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            content["plain_text"] = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        elif part.get_content_type() == "text/html":
                            content["html_content"] = part.get_payload(decode=True).decode('utf-8', errors='ignore')

            # Listar adjuntos
            attachments = []
            attachment_dir = self.attachments_dir / email_id
            if attachment_dir.exists():
                for file_path in attachment_dir.iterdir():
                    if file_path.is_file():
                        file_type_info = self.get_file_type_info(file_path.name)
                        attachments.append({
                            'name': file_path.name,
                            'size': file_path.stat().st_size,
                            'path': str(file_path),
                            'type': file_type_info['category'],
                            'extension': file_type_info['extension'],
                            'color_class': file_type_info['color_class'],
                            'mime_type': file_type_info['mime_type']
                        })

            # Obtener clasificación específica
            classification_info = None
            for email_info in self.data.get('emails', []):
                if email_info['email_id'] == email_id:
                    classification_info = email_info
                    break

            return {
                'metadata': metadata,
                'content': content,
                'attachments': attachments,
                'classification': classification_info
            }

        except Exception as e:
            return {'error': str(e)}

# Instancia global del dashboard
dashboard = EmailDashboard()

@app.route('/')
def index():
    """Página principal del dashboard"""
    stats = dashboard.get_summary_stats()
    charts = dashboard.create_charts()

    # Convertir gráficos a JSON para enviar al frontend
    charts_json = {}
    for name, chart in charts.items():
        charts_json[name] = plotly.utils.PlotlyJSONEncoder().encode({
            'data': chart['data'],
            'layout': chart['layout']
        })

    return render_template('dashboard.html', stats=stats, charts=charts_json)

@app.route('/search')
def search():
    """Página de búsqueda avanzada"""
    # Obtener opciones para filtros
    folders = dashboard.df['folder'].unique().tolist() if len(dashboard.df) > 0 else []
    classifications = ['cotizacion', 'renovacion', 'endoso', 'sin_clasificar']

    return render_template('search.html', folders=folders, classifications=classifications)

@app.route('/api/search')
def api_search():
    """API de búsqueda con paginación"""
    query = request.args.get('query', '')
    classification = request.args.get('classification', '')
    folder = request.args.get('folder', '')
    has_attachments = request.args.get('has_attachments')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 50))

    # Convertir has_attachments a boolean
    if has_attachments == 'true':
        has_attachments = True
    elif has_attachments == 'false':
        has_attachments = False
    else:
        has_attachments = None

    results = dashboard.search_emails(query, classification, folder, has_attachments, date_from, date_to, page, per_page)

    return jsonify(results)

@app.route('/email/<email_id>')
def view_email(email_id):
    """Ver email individual"""
    email_data = dashboard.get_email_content(email_id)

    if 'error' in email_data:
        abort(404)

    return render_template('email_detail.html',
                          email=email_data,
                          email_id=email_id,
                          GOOGLE_DRIVE_FOLDER_URL=GOOGLE_DRIVE_FOLDER_URL)

@app.route('/download/<email_id>')
def download_email_files(email_id):
    """Descargar todos los archivos de un email como ZIP"""
    try:
        # Crear ZIP en memoria
        memory_file = io.BytesIO()

        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            # Añadir archivo .eml
            eml_file = dashboard.emails_dir / f"{email_id}.eml"
            if eml_file.exists():
                zf.write(eml_file, f"{email_id}.eml")

            # Añadir metadatos
            metadata_file = dashboard.metadata_dir / f"{email_id}.json"
            if metadata_file.exists():
                zf.write(metadata_file, f"{email_id}_metadata.json")

            # Añadir adjuntos
            attachment_dir = dashboard.attachments_dir / email_id
            if attachment_dir.exists():
                for file_path in attachment_dir.iterdir():
                    if file_path.is_file():
                        zf.write(file_path, f"attachments/{file_path.name}")

        memory_file.seek(0)

        return send_file(
            io.BytesIO(memory_file.read()),
            mimetype='application/zip',
            as_attachment=True,
            download_name=f"{email_id}_complete.zip"
        )

    except Exception as e:
        abort(500)

@app.route('/download/attachment/<email_id>/<filename>')
def download_attachment(email_id, filename):
    """Descargar adjunto específico"""
    try:
        attachment_path = dashboard.attachments_dir / email_id / filename

        if not attachment_path.exists():
            abort(404)

        # Obtener información del tipo de archivo para envío correcto
        file_type_info = dashboard.get_file_type_info(filename)

        return send_file(
            attachment_path,
            as_attachment=True,
            download_name=filename,
            mimetype=file_type_info['mime_type']
        )

    except Exception as e:
        print(f"Error descargando adjunto {filename} del email {email_id}: {e}")
        abort(500)

@app.route('/api/stats')
def api_stats():
    """API para estadísticas en tiempo real"""
    stats = dashboard.get_summary_stats()
    return jsonify(stats)

if __name__ == '__main__':
    print("🌐 Iniciando Dashboard de Clasificación de Emails")
    print("📊 Accede a: http://localhost:3000")
    print("=" * 50)

    app.run(debug=True, host='0.0.0.0', port=3000)