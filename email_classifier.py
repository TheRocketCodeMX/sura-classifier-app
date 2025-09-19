#!/usr/bin/env python3
"""
Clasificador Automático de Correos de Seguros
Clasifica emails en: Cotización, Renovación, Endoso
Basado en especificaciones del negocio de seguros
"""

import os
import json
import re
import openpyxl
from pathlib import Path
from typing import Dict, List, Any
import PyPDF2
import email
from bs4 import BeautifulSoup


class EmailClassifier:
    def __init__(self, output_dir="output"):
        self.output_dir = Path(output_dir)
        self.metadata_dir = self.output_dir / "metadata"
        self.attachments_dir = self.output_dir / "attachments"
        self.classification_dir = self.output_dir / "classification"

        # Crear directorio de clasificación
        self.classification_dir.mkdir(exist_ok=True)

        # Patrones de palabras clave
        self.setup_patterns()

    def setup_patterns(self):
        """Configurar patrones de clasificación según especificaciones"""

        # COTIZACIÓN - Criterios de aceptación 1 y 2
        self.cotizacion_asunto_patterns = [
            r'\bCOTIZACI[OÓ]N\b',
            r'\bCOT\.\b',
            r'\bCOT\b(?!\w)',  # COT pero no como parte de otra palabra
            r'\bCOT RESIDENCIAL\b',
            r'\bAPOYO COTIZACION\b',
            r'\bAGENTE\s+\d+',
            r'\bAG\s+\d+'
        ]

        # COTIZACIÓN - Criterio de aceptación 3 (cuerpo del mensaje)
        self.cotizacion_cuerpo_patterns = [
            r'solicito su apoyo cotizando',
            r'apoyo para cotizar',
            r'solicitud de cotizaci[óo]n',
            r'\bAGENTE\s+\d+'
        ]

        # RENOVACIÓN - Criterios de aceptación 1, 2 y 3
        self.renovacion_asunto_patterns = [
            r'\bRENOVACI[OÓ]N\b',
            r'\bRENOVAR\b',
            r'\bRENOVACIONES\b',
            r'\bREHABILITACI[OÓ]N\b',
            r'\bPR[OÓ]RROGA\b',
            r'\bRV\b',
            r'\bRENOV\b',
            r'\bCOTI RENOVACI[OÓ]N\b'
        ]

        # RENOVACIÓN - Criterios de aceptación 4, 5 y 6 (cuerpo)
        self.renovacion_cuerpo_patterns = [
            r'vigencia pr[óo]xima a vencer',
            r'solicito renovaci[óo]n',
            r'renovar p[óo]liza',
            r'dar continuidad a la p[óo]liza',
            r'pr[óo]rroga de la vigencia',
            r'rehabilitaci[óo]n de p[óo]liza',
            r'continuidad de p[óo]liza',
            r'renovar p[óo]liza \d+'
        ]

        # ENDOSO - Criterios de aceptación 1, 2 y 3 (MEJORADO)
        self.endoso_asunto_patterns = [
            r'\bENDOSO\b',
            r'\bENDOSOS\b',
            r'\bENDOSAR\b',
            r'\bENDORSEMENT\b',
            r'\bENDOSO [AB]\b',
            r'\bENDOSO DE BP\b',
            r'\bENDOSO ESPECIAL\b',
            r'\bENDOSO.*MODIFICACI[OÓ]N\b',  # Para casos como "Endoso (Modificación)"
            r'\bMODIFICACI[OÓ]N.*ENDOSO\b',  # Orden inverso
            r'\bINCISO \d+\b',
            r'CORRECCI[OÓ]N DE DATO',
            r'CAMBIO DE COBERTURA',
            r'INCREMENTO DE SUMA ASEGURADA',
            r'\bOT-\d+',  # Para números de operación como "OT-0710313"
            r'DOCUMENTO \d+',  # Para referencias a documentos
            r'MODIFICACI[OÓ]N.*NO\s+OT',  # Patrones específicos del ejemplo
            r'ENDOSO.*\(.*(MODIFICACI[OÓ]N|CAMBIO|CORRECCI[OÓ]N).*\)'  # Endoso con paréntesis
        ]

        # ENDOSO - Criterios de aceptación 4, 5 y 6 (cuerpo)
        self.endoso_cuerpo_patterns = [
            r'correcci[óo]n de dato',
            r'modificar inciso',
            r'incluir.{0,20}beneficiario',
            r'excluir.{0,20}beneficiario',
            r'cambiar suma asegurada',
            r'actualizaci[óo]n de cobertura',
            r'actualizaci[óo]n de beneficios',
            r'incorporaci[óo]n de cl[áa]usulas'
        ]

    def extract_agente_code(self, text: str) -> str:
        """Extraer código de agente del texto"""
        patterns = [
            r'\bAGENTE\s+(\d+)',
            r'\bAG\s+(\d+)'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)
        return ""

    def extract_poliza_number(self, text: str) -> str:
        """Extraer número de póliza del texto"""
        patterns = [
            r'\bp[óo]liza\s+(\d+)',
            r'\bP[ÓO]LIZA\s+(\d+)',
            r'\bOT\s+(\d+)',
            r'\bN[ÚU]MERO\s+(\d+)'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)
        return ""

    def extract_email_content(self, email_id: str) -> Dict[str, str]:
        """Extraer contenido del email (.eml file)"""
        emails_dir = self.output_dir / "emails"
        eml_file = emails_dir / f"{email_id}.eml"

        content = {
            'plain_text': '',
            'html_content': '',
            'combined_text': ''
        }

        if not eml_file.exists():
            return content

        try:
            with open(eml_file, 'r', encoding='utf-8') as f:
                msg = email.message_from_file(f)

                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            content['plain_text'] = payload.decode('utf-8', errors='ignore')
                    elif part.get_content_type() == "text/html":
                        payload = part.get_payload(decode=True)
                        if payload:
                            html_content = payload.decode('utf-8', errors='ignore')
                            content['html_content'] = html_content

                            # Extraer texto del HTML usando BeautifulSoup
                            try:
                                soup = BeautifulSoup(html_content, 'html.parser')
                                content['combined_text'] = soup.get_text(separator=' ', strip=True)
                            except:
                                content['combined_text'] = content['plain_text']

                # Si no hay texto plano, usar el extraído del HTML
                if not content['plain_text'] and content['combined_text']:
                    content['plain_text'] = content['combined_text']

        except Exception as e:
            print(f"Error extrayendo contenido del email {email_id}: {e}")

        return content

    def analyze_attachments(self, email_id: str) -> Dict[str, Any]:
        """Analizar adjuntos según criterios de aceptación"""
        attachment_dir = self.attachments_dir / email_id
        attachment_info = {
            'has_slip': False,
            'slip_complete': False,
            'slip_files': [],
            'pdf_cotizacion': [],
            'pdf_poliza': [],
            'pdf_renovacion': [],
            'pdf_endoso': [],
            'excel_files': [],
            'total_attachments': 0
        }

        if not attachment_dir.exists():
            return attachment_info

        for file_path in attachment_dir.iterdir():
            if file_path.is_file():
                attachment_info['total_attachments'] += 1
                filename = file_path.name.upper()

                # Criterio de aceptación 4 - Archivos SLIP
                if filename.endswith(('.XLSX', '.XLS')) and 'SLIP' in filename:
                    attachment_info['has_slip'] = True
                    attachment_info['slip_files'].append(file_path.name)

                    # Criterio de aceptación 5 - Verificar si el SLIP está completo
                    try:
                        wb = openpyxl.load_workbook(file_path, data_only=True)
                        ws = wb.active
                        filled_cells = 0

                        for row in ws.iter_rows():
                            for cell in row:
                                if cell.value is not None and str(cell.value).strip():
                                    filled_cells += 1

                        # Si tiene más de 5 celdas con datos, se considera completo
                        attachment_info['slip_complete'] = filled_cells > 5
                        wb.close()
                    except:
                        pass

                # Excel files en general
                elif filename.endswith(('.XLSX', '.XLS')):
                    attachment_info['excel_files'].append(file_path.name)

                # Criterio de aceptación 6 - PDFs de cotización
                elif filename.endswith('.PDF') and any(word in filename for word in
                    ['COTIZACION', 'COT', 'PROPUESTA', 'SEGURO MULTIPLE EMPRESARIAL']):
                    attachment_info['pdf_cotizacion'].append(file_path.name)

                # Criterio de aceptación 7 - PDFs de póliza
                elif filename.endswith('.PDF') and any(word in filename for word in
                    ['POLIZA', 'PBE', 'RECIBO']):
                    attachment_info['pdf_poliza'].append(file_path.name)

                # CA-Adjuntos 1 - PDFs de renovación
                elif filename.endswith('.PDF') and any(word in filename for word in
                    ['RENOVACION', 'RENOVAR', 'CONDICIONES DE RENOVACION', 'PRORROGA', 'REHABILITACION']):
                    attachment_info['pdf_renovacion'].append(file_path.name)

                # Criterio de aceptación 7 (Endoso) - PDFs de endoso
                elif filename.endswith('.PDF') and any(word in filename for word in
                    ['ENDOSO', 'BENEFICIOS', 'INCISO', 'CORRECCION', 'MODIFICACION']):
                    attachment_info['pdf_endoso'].append(file_path.name)

        return attachment_info

    def classify_cotizacion(self, metadata: Dict[str, Any], attachment_info: Dict[str, Any], email_content: Dict[str, str]) -> Dict[str, Any]:
        """Clasificar email como cotización según criterios de aceptación"""
        classification = {
            'is_cotizacion': False,
            'confidence': 0,
            'criteria_met': [],
            'agente_code': '',
            'client_name': '',
            'status': '',
            'details': {}
        }

        asunto = metadata.get('subject', '').upper()
        cuerpo = email_content.get('combined_text', '').upper()

        score = 0

        # Criterio de aceptación 1 - Palabras clave en asunto
        for pattern in self.cotizacion_asunto_patterns:
            if re.search(pattern, asunto, re.IGNORECASE):
                score += 30
                classification['criteria_met'].append('CA1: Palabra clave en asunto')
                break

        # Criterio de aceptación 2 - Código de agente en asunto o cuerpo
        agente_code = self.extract_agente_code(asunto) or self.extract_agente_code(cuerpo)
        if agente_code:
            classification['agente_code'] = agente_code
            score += 20
            classification['criteria_met'].append('CA2: Código de agente detectado')

        # Criterio de aceptación 3 - Palabras clave en cuerpo del mensaje
        for pattern in self.cotizacion_cuerpo_patterns:
            if re.search(pattern, cuerpo, re.IGNORECASE):
                score += 15
                classification['criteria_met'].append('CA3: Palabra clave en cuerpo del email')
                break

        # Criterio de aceptación 4 - SLIP presente
        if attachment_info['has_slip']:
            score += 25
            classification['criteria_met'].append('CA4: Archivo SLIP detectado')

            # Criterio de aceptación 5 - SLIP completo o vacío
            if attachment_info['slip_complete']:
                classification['status'] = 'Cotización con información completa'
                score += 15
                classification['criteria_met'].append('CA5: SLIP completo')
            else:
                classification['status'] = 'Cotización pendiente'
                classification['criteria_met'].append('CA5: SLIP vacío o incompleto')

        # Criterio de aceptación 6 - PDFs de cotización
        if attachment_info['pdf_cotizacion']:
            score += 20
            classification['criteria_met'].append('CA6: PDF de cotización presente')

        # Criterio de aceptación 7 y 10 - Excluir pólizas vigentes
        if attachment_info['pdf_poliza']:
            score -= 15
            classification['criteria_met'].append('CA7/CA10: Contiene documentos de póliza vigente')

        classification['confidence'] = min(100, score)
        classification['is_cotizacion'] = score >= 40

        if not classification['status'] and classification['is_cotizacion']:
            classification['status'] = 'Cotización detectada'

        classification['details'] = {
            'slip_files': attachment_info['slip_files'],
            'pdf_files': attachment_info['pdf_cotizacion'],
            'total_attachments': attachment_info['total_attachments']
        }

        return classification

    def classify_renovacion(self, metadata: Dict[str, Any], attachment_info: Dict[str, Any], email_content: Dict[str, str]) -> Dict[str, Any]:
        """Clasificar email como renovación según criterios de aceptación"""
        classification = {
            'is_renovacion': False,
            'confidence': 0,
            'criteria_met': [],
            'poliza_number': '',
            'client_name': '',
            'status': '',
            'details': {}
        }

        asunto = metadata.get('subject', '').upper()
        cuerpo = email_content.get('combined_text', '').upper()
        score = 0

        # Criterio de aceptación 1 - Palabras clave en asunto
        for pattern in self.renovacion_asunto_patterns:
            if re.search(pattern, asunto, re.IGNORECASE):
                score += 35
                classification['criteria_met'].append('CA1: Palabra clave de renovación en asunto')
                break

        # Criterio de aceptación 2 - Número de póliza en asunto o cuerpo
        poliza_number = self.extract_poliza_number(asunto) or self.extract_poliza_number(cuerpo)
        if poliza_number:
            classification['poliza_number'] = poliza_number
            score += 20
            classification['criteria_met'].append('CA2: Número de póliza detectado')

        # Criterios de aceptación 4, 5 y 6 - Palabras clave en cuerpo
        for pattern in self.renovacion_cuerpo_patterns:
            if re.search(pattern, cuerpo, re.IGNORECASE):
                score += 15
                classification['criteria_met'].append('CA4-6: Palabra clave de renovación en cuerpo')
                break

        # CA-Adjuntos 1 - PDFs de renovación
        if attachment_info['pdf_renovacion']:
            score += 25
            classification['criteria_met'].append('CA-Adj1: PDF de renovación presente')

        # CA-Adjuntos 2 - Documentos de póliza con mención de renovación
        if attachment_info['pdf_poliza'] and any(word in asunto for word in ['RENOVACION', 'RENOVAR']):
            score += 20
            classification['criteria_met'].append('CA-Adj2: Documentos de póliza + renovación')

        # CA-Adjuntos 5 - Prevenir falsos positivos con cotización
        if attachment_info['pdf_cotizacion'] and 'RENOVACION' not in asunto:
            score -= 20
            classification['criteria_met'].append('CA-Adj5: Prevención falso positivo cotización')

        classification['confidence'] = min(100, score)
        classification['is_renovacion'] = score >= 40

        # Criterio de aceptación 10 - Determinar estado
        if classification['is_renovacion']:
            if attachment_info['total_attachments'] > 0 and poliza_number:
                classification['status'] = 'Renovación con información completa'
            else:
                classification['status'] = 'Renovación pendiente'

        classification['details'] = {
            'pdf_renovacion': attachment_info['pdf_renovacion'],
            'pdf_poliza': attachment_info['pdf_poliza'],
            'total_attachments': attachment_info['total_attachments']
        }

        return classification

    def classify_endoso(self, metadata: Dict[str, Any], attachment_info: Dict[str, Any], email_content: Dict[str, str]) -> Dict[str, Any]:
        """Clasificar email como endoso según criterios de aceptación"""
        classification = {
            'is_endoso': False,
            'confidence': 0,
            'criteria_met': [],
            'poliza_number': '',
            'endoso_type': '',
            'client_name': '',
            'status': '',
            'details': {}
        }

        asunto = metadata.get('subject', '').upper()
        cuerpo = email_content.get('combined_text', '').upper()
        score = 0

        # Criterio de aceptación 1 y 2 - Palabras clave en asunto
        for pattern in self.endoso_asunto_patterns:
            match = re.search(pattern, asunto, re.IGNORECASE)
            if match:
                score += 35
                classification['criteria_met'].append('CA1-2: Palabra clave de endoso en asunto')

                # Detectar tipo de endoso específico
                if 'ENDOSO A' in asunto:
                    classification['endoso_type'] = 'A'
                elif 'ENDOSO B' in asunto:
                    classification['endoso_type'] = 'B'
                elif 'ENDOSO DE BP' in asunto:
                    classification['endoso_type'] = 'BP'
                elif 'ENDOSO ESPECIAL' in asunto:
                    classification['endoso_type'] = 'ESPECIAL'
                break

        # Criterio de aceptación 3 y 4 - Palabras clave en cuerpo del mensaje
        for pattern in self.endoso_cuerpo_patterns:
            if re.search(pattern, cuerpo, re.IGNORECASE):
                score += 15
                classification['criteria_met'].append('CA3-4: Palabra clave de endoso en cuerpo')
                break

        # Criterio de aceptación 5 - Referencia a póliza vigente
        poliza_number = self.extract_poliza_number(asunto) or self.extract_poliza_number(cuerpo)
        if poliza_number:
            classification['poliza_number'] = poliza_number
            score += 25
            classification['criteria_met'].append('CA5: Referencia a póliza vigente')

        # Criterio de aceptación 7 - PDFs de endoso
        if attachment_info['pdf_endoso']:
            score += 20
            classification['criteria_met'].append('CA7: PDF de endoso presente')

        # Criterio de aceptación 8 - Documentos de respaldo
        if attachment_info['pdf_poliza'] and any(word in asunto for word in ['ENDOSO', 'MODIFICACION', 'CORRECCION']):
            score += 15
            classification['criteria_met'].append('CA8: Documentos de respaldo de endoso')

        classification['confidence'] = min(100, score)
        classification['is_endoso'] = score >= 30

        # Criterio de aceptación 10 - Determinar completitud
        if classification['is_endoso']:
            if poliza_number and (attachment_info['pdf_endoso'] or attachment_info['pdf_poliza']):
                classification['status'] = 'Endoso completo'
            else:
                classification['status'] = 'Endoso incompleto'

        classification['details'] = {
            'pdf_endoso': attachment_info['pdf_endoso'],
            'excel_files': attachment_info['excel_files'],
            'total_attachments': attachment_info['total_attachments']
        }

        return classification

    def classify_email(self, email_id: str) -> Dict[str, Any]:
        """Clasificar un email individual"""
        # Cargar metadatos
        metadata_file = self.metadata_dir / f"{email_id}.json"
        if not metadata_file.exists():
            return {'error': f'Metadatos no encontrados para {email_id}'}

        with open(metadata_file, 'r', encoding='utf-8') as f:
            metadata = json.load(f)

        # Extraer contenido del email
        email_content = self.extract_email_content(email_id)

        # Analizar adjuntos
        attachment_info = self.analyze_attachments(email_id)

        # Realizar clasificaciones
        cotizacion = self.classify_cotizacion(metadata, attachment_info, email_content)
        renovacion = self.classify_renovacion(metadata, attachment_info, email_content)
        endoso = self.classify_endoso(metadata, attachment_info, email_content)

        # Determinar clasificación principal
        classifications = [
            ('cotizacion', cotizacion),
            ('renovacion', renovacion),
            ('endoso', endoso)
        ]

        # Ordenar por confianza
        classifications.sort(key=lambda x: x[1]['confidence'], reverse=True)

        primary_class = classifications[0]

        result = {
            'email_id': email_id,
            'metadata': metadata,
            'attachment_analysis': attachment_info,
            'classifications': {
                'cotizacion': cotizacion,
                'renovacion': renovacion,
                'endoso': endoso
            },
            'primary_classification': {
                'type': primary_class[0] if primary_class[1]['confidence'] >= 30 else 'sin_clasificar',
                'confidence': primary_class[1]['confidence'],
                'status': primary_class[1].get('status', ''),
                'details': primary_class[1]
            }
        }

        return result

    def classify_all_emails(self) -> Dict[str, Any]:
        """Clasificar todos los emails procesados"""
        results = {
            'total_emails': 0,
            'cotizacion': 0,
            'renovacion': 0,
            'endoso': 0,
            'sin_clasificar': 0,
            'emails': []
        }

        if not self.metadata_dir.exists():
            return {'error': 'Directorio de metadatos no encontrado'}

        for metadata_file in self.metadata_dir.glob('*.json'):
            if metadata_file.name == 'progress.json':
                continue

            email_id = metadata_file.stem
            classification = self.classify_email(email_id)

            if 'error' not in classification:
                results['emails'].append(classification)
                results['total_emails'] += 1

                # Contar por categoría
                primary_type = classification['primary_classification']['type']
                if primary_type in results:
                    results[primary_type] += 1

        # Guardar resultados
        results_file = self.classification_dir / 'classification_results.json'
        with open(results_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        return results

    def generate_report(self) -> str:
        """Generar reporte de clasificación"""
        results = self.classify_all_emails()

        if 'error' in results:
            return f"Error: {results['error']}"

        report = f"""
REPORTE DE CLASIFICACIÓN AUTOMÁTICA DE CORREOS
==============================================

📊 RESUMEN GENERAL:
- Total de emails procesados: {results['total_emails']}
- Cotizaciones: {results['cotizacion']} ({results['cotizacion']/results['total_emails']*100:.1f}%)
- Renovaciones: {results['renovacion']} ({results['renovacion']/results['total_emails']*100:.1f}%)
- Endosos: {results['endoso']} ({results['endoso']/results['total_emails']*100:.1f}%)
- Sin clasificar: {results['sin_clasificar']} ({results['sin_clasificar']/results['total_emails']*100:.1f}%)

📋 DETALLE POR CATEGORÍA:

COTIZACIONES ({results['cotizacion']} emails):
"""

        # Añadir detalles de cotizaciones
        cotizaciones = [e for e in results['emails'] if e['primary_classification']['type'] == 'cotizacion']
        for email in cotizaciones[:5]:  # Mostrar solo los primeros 5
            report += f"  - {email['email_id']}: {email['metadata'].get('subject', 'Sin asunto')[:60]}...\n"

        if len(cotizaciones) > 5:
            report += f"  ... y {len(cotizaciones) - 5} más\n"

        report += f"\nRENOVACIONES ({results['renovacion']} emails):\n"
        renovaciones = [e for e in results['emails'] if e['primary_classification']['type'] == 'renovacion']
        for email in renovaciones[:5]:
            report += f"  - {email['email_id']}: {email['metadata'].get('subject', 'Sin asunto')[:60]}...\n"

        if len(renovaciones) > 5:
            report += f"  ... y {len(renovaciones) - 5} más\n"

        report += f"\nENDOSOS ({results['endoso']} emails):\n"
        endosos = [e for e in results['emails'] if e['primary_classification']['type'] == 'endoso']
        for email in endosos[:5]:
            report += f"  - {email['email_id']}: {email['metadata'].get('subject', 'Sin asunto')[:60]}...\n"

        if len(endosos) > 5:
            report += f"  ... y {len(endosos) - 5} más\n"

        report += f"\n💾 Resultados detallados guardados en: {self.classification_dir}/classification_results.json\n"

        return report


def main():
    """Función principal"""
    print("🔍 CLASIFICADOR AUTOMÁTICO DE CORREOS DE SEGUROS")
    print("=" * 60)

    classifier = EmailClassifier()

    print("Iniciando clasificación de todos los emails...")
    report = classifier.generate_report()

    print(report)

    # Guardar reporte
    report_file = classifier.classification_dir / 'classification_report.txt'
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(report)

    print(f"\n📄 Reporte guardado en: {report_file}")


if __name__ == "__main__":
    main()