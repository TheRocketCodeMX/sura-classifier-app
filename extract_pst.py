#!/usr/bin/env python3
"""
Script principal para extraer emails de archivos PST
Uso: python extract_pst.py <archivo_pst>
"""

import sys
import os
from pathlib import Path
from pst_extractor import PSTExtractor


def main():
    print("=" * 60)
    print("PST Email Extractor")
    print("=" * 60)

    if len(sys.argv) < 2:
        print("\nUso:")
        print("  python extract_pst.py <archivo_pst>")
        print("\nEjemplo:")
        print("  python extract_pst.py /ruta/al/archivo.pst")
        print("\nEl script creará la siguiente estructura de salida:")
        print("  output/")
        print("  ├── emails/          # Archivos .eml individuales")
        print("  ├── attachments/     # Adjuntos organizados por email")
        print("  ├── metadata/        # JSON con metadatos de cada email")
        print("  └── progress.json    # Estado del procesamiento")
        return

    pst_file = sys.argv[1]

    # Verificar que el archivo existe
    if not os.path.exists(pst_file):
        print(f"\nError: El archivo '{pst_file}' no existe")
        return

    # Verificar que es un archivo PST
    if not pst_file.lower().endswith('.pst'):
        print(f"\nAdvertencia: El archivo '{pst_file}' no tiene extensión .pst")
        response = input("¿Continuar de todos modos? (s/N): ")
        if response.lower() not in ['s', 'si', 'sí', 'y', 'yes']:
            return

    # Mostrar información del archivo
    file_size = os.path.getsize(pst_file)
    print(f"\nArchivo a procesar: {pst_file}")
    print(f"Tamaño: {file_size / (1024**3):.2f} GB")

    # Confirmar antes de proceder
    print(f"\nSe creará la estructura de salida en: {Path.cwd() / 'output'}")
    response = input("¿Proceder con la extracción? (S/n): ")
    if response.lower() in ['n', 'no']:
        print("Extracción cancelada")
        return

    try:
        # Crear extractor y ejecutar
        print("\nInicializando extractor...")
        extractor = PSTExtractor(pst_file)
        extractor.extract()

        print(f"\n" + "=" * 60)
        print("EXTRACCIÓN COMPLETADA")
        print("=" * 60)
        print(f"Total de emails procesados: {extractor.processed_count}")
        print(f"Archivos guardados en: {extractor.output_dir}")
        print("\nEstructura creada:")
        print(f"  📧 Emails: {extractor.emails_dir}")
        print(f"  📎 Adjuntos: {extractor.attachments_dir}")
        print(f"  📋 Metadatos: {extractor.metadata_dir}")
        print(f"  📊 Progreso: {extractor.progress_file}")

    except KeyboardInterrupt:
        print("\n\nExtracción interrumpida por el usuario")
        print("El progreso se ha guardado en progress.json")
        print("Puede reanudar la extracción ejecutando el mismo comando")

    except Exception as e:
        print(f"\n\nError durante la extracción: {e}")
        print("Revise el archivo PST y vuelva a intentar")


if __name__ == "__main__":
    main()