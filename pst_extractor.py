#!/usr/bin/env python3
"""
PST Email Extractor
Extrae emails, adjuntos y metadatos de archivos PST de Outlook
"""

import os
import json
import base64
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from pathlib import Path
import pypff


class PSTExtractor:
    def __init__(self, pst_file_path, output_dir="output"):
        self.pst_file_path = pst_file_path
        self.output_dir = Path(output_dir)
        self.emails_dir = self.output_dir / "emails"
        self.attachments_dir = self.output_dir / "attachments"
        self.metadata_dir = self.output_dir / "metadata"
        self.progress_file = self.output_dir / "progress.json"

        # Crear directorios si no existen
        for directory in [self.emails_dir, self.attachments_dir, self.metadata_dir]:
            directory.mkdir(parents=True, exist_ok=True)

        self.processed_count = 0
        self.total_count = 0
        self.progress_data = self.load_progress()

    def load_progress(self):
        """Carga el progreso previo si existe"""
        if self.progress_file.exists():
            try:
                with open(self.progress_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {
            'processed_emails': [],
            'total_processed': 0,
            'start_time': None,
            'last_update': None
        }

    def save_progress(self):
        """Guarda el progreso actual"""
        self.progress_data['total_processed'] = self.processed_count
        self.progress_data['last_update'] = datetime.now().isoformat()

        with open(self.progress_file, 'w', encoding='utf-8') as f:
            json.dump(self.progress_data, f, indent=2, ensure_ascii=False)

    def extract_email_content(self, message):
        """Extrae el contenido de texto del email"""
        plain_text = ""
        html_content = ""

        try:
            # Obtener contenido en texto plano
            if hasattr(message, 'plain_text_body') and message.plain_text_body:
                plain_text = message.plain_text_body
                # Convertir bytes a string si es necesario
                if isinstance(plain_text, bytes):
                    plain_text = plain_text.decode('utf-8', errors='ignore')

            # Obtener contenido HTML
            if hasattr(message, 'html_body') and message.html_body:
                html_content = message.html_body
                # Convertir bytes a string si es necesario
                if isinstance(html_content, bytes):
                    html_content = html_content.decode('utf-8', errors='ignore')
        except:
            pass

        return plain_text, html_content

    def extract_attachments(self, message, email_id):
        """Extrae adjuntos del email"""
        attachments = []

        try:
            # Intentar diferentes métodos para obtener adjuntos con manejo de errores
            num_attachments = 0
            try:
                if hasattr(message, 'get_number_of_attachments'):
                    num_attachments = message.get_number_of_attachments()
                elif hasattr(message, 'number_of_attachments'):
                    num_attachments = message.number_of_attachments
            except Exception as e:
                # Si hay error obteniendo el número de adjuntos, continuar sin adjuntos
                return []

            for attachment_index in range(num_attachments):
                attachment = message.get_attachment(attachment_index)

                if attachment:
                    # Crear directorio para adjuntos de este email
                    email_attachments_dir = self.attachments_dir / email_id
                    email_attachments_dir.mkdir(exist_ok=True)

                    # Obtener nombre del adjunto
                    filename = "attachment_{}".format(attachment_index)
                    if hasattr(attachment, 'name') and attachment.name:
                        filename = attachment.name
                    elif hasattr(attachment, 'long_filename') and attachment.long_filename:
                        filename = attachment.long_filename

                    # Guardar adjunto
                    attachment_path = email_attachments_dir / filename

                    try:
                        # Leer datos del adjunto - usar get_size() en lugar de size
                        attachment_size = attachment.get_size()
                        attachment_data = attachment.read_buffer(attachment_size)

                        with open(attachment_path, 'wb') as f:
                            f.write(attachment_data)

                        attachments.append({
                            'filename': filename,
                            'size': attachment_size,
                            'path': str(attachment_path.relative_to(self.output_dir))
                        })
                    except Exception as e:
                        print(f"Error extrayendo adjunto {filename}: {e}")
        except Exception as e:
            print(f"Error procesando adjuntos: {e}")

        return attachments

    def create_eml_file(self, message, email_id, plain_text, html_content):
        """Crea un archivo .eml estándar"""
        try:
            # Crear objeto email
            msg = MIMEMultipart('alternative')

            # Headers básicos
            if hasattr(message, 'subject') and message.subject:
                msg['Subject'] = message.subject

            if hasattr(message, 'sender_name') and message.sender_name:
                sender = message.sender_name
                if hasattr(message, 'sender_email_address') and message.sender_email_address:
                    sender = f"{message.sender_name} <{message.sender_email_address}>"
                msg['From'] = sender

            if hasattr(message, 'delivery_time') and message.delivery_time:
                msg['Date'] = message.delivery_time.strftime("%a, %d %b %Y %H:%M:%S %z")

            # Contenido
            if plain_text:
                msg.attach(MIMEText(plain_text, 'plain'))
            if html_content:
                msg.attach(MIMEText(html_content, 'html'))

            # Guardar archivo .eml
            eml_path = self.emails_dir / f"{email_id}.eml"
            with open(eml_path, 'w', encoding='utf-8') as f:
                f.write(str(msg))

            return str(eml_path.relative_to(self.output_dir))

        except Exception as e:
            print(f"Error creando archivo EML para {email_id}: {e}")
            return None

    def extract_metadata(self, message, email_id, folder_name):
        """Extrae metadatos del email"""
        metadata = {
            'id': email_id,
            'folder': folder_name,
            'extraction_date': datetime.now().isoformat()
        }

        # Información básica
        try:
            if hasattr(message, 'subject'):
                metadata['subject'] = message.subject
            if hasattr(message, 'sender_name'):
                metadata['sender_name'] = message.sender_name
            if hasattr(message, 'sender_email_address'):
                metadata['sender_email'] = message.sender_email_address
            if hasattr(message, 'delivery_time'):
                metadata['delivery_time'] = message.delivery_time.isoformat() if message.delivery_time else None
            if hasattr(message, 'creation_time'):
                metadata['creation_time'] = message.creation_time.isoformat() if message.creation_time else None
            if hasattr(message, 'modification_time'):
                metadata['modification_time'] = message.modification_time.isoformat() if message.modification_time else None

            # Información de tamaño
            try:
                metadata['size'] = message.get_size()
            except:
                pass

            # Número de adjuntos
            try:
                metadata['attachment_count'] = message.get_number_of_attachments()
            except:
                pass

        except Exception as e:
            print(f"Error extrayendo metadatos básicos: {e}")

        return metadata

    def process_folder(self, folder, folder_path=""):
        """Procesa una carpeta y sus subcarpetas recursivamente"""
        folder_name = folder_path
        if hasattr(folder, 'name') and folder.name:
            folder_name = f"{folder_path}/{folder.name}" if folder_path else folder.name

        print(f"Procesando carpeta: {folder_name}")

        # Procesar mensajes en esta carpeta
        try:
            # Usar los métodos correctos: get_number_of_sub_messages y get_sub_message
            num_messages = 0
            if hasattr(folder, 'get_number_of_sub_messages'):
                num_messages = folder.get_number_of_sub_messages()
            elif hasattr(folder, 'number_of_sub_messages'):
                num_messages = folder.number_of_sub_messages

            if num_messages > 0:
                print(f"  📧 {num_messages} mensajes encontrados")
                for message_index in range(num_messages):
                    message = folder.get_sub_message(message_index)
                    if message:
                        self.process_message(message, folder_name)
            else:
                print(f"  📭 Sin mensajes en esta carpeta")
        except Exception as e:
            print(f"  ❌ Error procesando mensajes en {folder_name}: {e}")
            # Intentar método alternativo con iteración directa usando get_sub_message
            try:
                message_index = 0
                while True:
                    message = folder.get_sub_message(message_index)
                    if message:
                        self.process_message(message, folder_name)
                        message_index += 1
                    else:
                        break
                if message_index > 0:
                    print(f"  ✅ {message_index} mensajes procesados por iteración directa")
            except:
                pass

        # Procesar subcarpetas
        try:
            # Intentar diferentes métodos para obtener el número de subcarpetas
            num_subfolders = 0
            if hasattr(folder, 'get_number_of_sub_folders'):
                num_subfolders = folder.get_number_of_sub_folders()
            elif hasattr(folder, 'number_of_sub_folders'):
                num_subfolders = folder.number_of_sub_folders

            for subfolder_index in range(num_subfolders):
                subfolder = folder.get_sub_folder(subfolder_index)
                if subfolder:
                    self.process_folder(subfolder, folder_name)
        except Exception as e:
            print(f"  ❌ Error procesando subcarpetas en {folder_name}: {e}")

    def process_message(self, message, folder_name):
        """Procesa un mensaje individual"""
        self.processed_count += 1
        email_id = f"email_{self.processed_count:06d}"

        # Verificar si ya fue procesado
        if email_id in self.progress_data.get('processed_emails', []):
            return

        try:
            # Extraer contenido
            plain_text, html_content = self.extract_email_content(message)

            # Extraer adjuntos
            attachments = self.extract_attachments(message, email_id)

            # Crear archivo .eml
            eml_path = self.create_eml_file(message, email_id, plain_text, html_content)

            # Extraer metadatos
            metadata = self.extract_metadata(message, email_id, folder_name)
            metadata['eml_file'] = eml_path
            metadata['attachments'] = attachments
            metadata['plain_text_length'] = len(plain_text) if plain_text else 0
            metadata['html_content_length'] = len(html_content) if html_content else 0

            # Guardar metadatos
            metadata_path = self.metadata_dir / f"{email_id}.json"
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)

            # Actualizar progreso
            self.progress_data['processed_emails'].append(email_id)

            # Mostrar progreso cada 10 emails
            if self.processed_count % 10 == 0:
                print(f"Procesados: {self.processed_count} emails")
                self.save_progress()

        except Exception as e:
            print(f"Error procesando email {email_id}: {e}")

    def extract(self):
        """Función principal de extracción"""
        try:
            print(f"Abriendo archivo PST: {self.pst_file_path}")

            # Abrir archivo PST
            pst_file = pypff.file()
            pst_file.open(self.pst_file_path)

            # Inicializar progreso
            if not self.progress_data['start_time']:
                self.progress_data['start_time'] = datetime.now().isoformat()

            print("Iniciando extracción...")

            # Obtener carpeta raíz
            root_folder = pst_file.get_root_folder()

            # Procesar todas las carpetas
            self.process_folder(root_folder)

            # Guardar progreso final
            self.save_progress()

            print(f"\nExtracción completada!")
            print(f"Total de emails procesados: {self.processed_count}")
            print(f"Archivos guardados en: {self.output_dir}")

            pst_file.close()

        except Exception as e:
            print(f"Error durante la extracción: {e}")
            self.save_progress()
            raise


def main():
    """Función principal"""
    import sys

    if len(sys.argv) < 2:
        print("Uso: python pst_extractor.py <archivo_pst>")
        print("Ejemplo: python pst_extractor.py /ruta/al/archivo.pst")
        sys.exit(1)

    pst_file_path = sys.argv[1]

    if not os.path.exists(pst_file_path):
        print(f"Error: El archivo {pst_file_path} no existe")
        sys.exit(1)

    # Crear extractor y ejecutar
    extractor = PSTExtractor(pst_file_path)
    extractor.extract()


if __name__ == "__main__":
    main()