# Configuración para Google Drive
# ==============================

# URL de la carpeta de Google Drive donde están los archivos .eml y adjuntos
# Reemplaza con tu URL real de Google Drive
GOOGLE_DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1cn75U0L6twGqmrfx6BmpHlQdDx-hBzO8?usp=drive_link"

# Instrucciones:
# 1. Sube la carpeta output/emails/ a Google Drive
# 2. Sube la carpeta output/attachments/ a Google Drive
# 3. Haz las carpetas públicas o compartidas
# 4. Copia la URL de la carpeta principal
# 5. Reemplaza la URL arriba
#
# Ejemplo de URL:
# "https://drive.google.com/drive/folders/1aBcDeFgHiJkLmNoPqRsTuVwXyZ123456"

# Configuración de archivos
EMAIL_FILE_PATTERN = "{email_id}.eml"  # Patrón del nombre de archivo
ATTACHMENT_FOLDER_PATTERN = "attachments/{email_id}/"  # Patrón de carpeta de adjuntos