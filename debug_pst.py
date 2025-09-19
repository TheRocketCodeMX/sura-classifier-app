#!/usr/bin/env python3
"""
Script de debug para explorar la estructura del PST
"""

import pypff
import sys

def explore_folder(folder, level=0, max_level=3):
    """Explora recursivamente la estructura de carpetas"""
    indent = "  " * level

    # Obtener nombre de la carpeta
    folder_name = "ROOT"
    if hasattr(folder, 'name') and folder.name:
        folder_name = folder.name

    print(f"{indent}📁 {folder_name}")

    # Intentar obtener número de mensajes
    num_messages = 0
    try:
        if hasattr(folder, 'get_number_of_messages'):
            num_messages = folder.get_number_of_messages()
        elif hasattr(folder, 'number_of_messages'):
            num_messages = folder.number_of_messages

        if num_messages > 0:
            print(f"{indent}  📧 {num_messages} mensajes")
    except Exception as e:
        print(f"{indent}  ❌ Error contando mensajes: {e}")

    # Explorar subcarpetas solo si no hemos llegado al máximo nivel
    if level < max_level:
        try:
            num_subfolders = 0
            if hasattr(folder, 'get_number_of_sub_folders'):
                num_subfolders = folder.get_number_of_sub_folders()
            elif hasattr(folder, 'number_of_sub_folders'):
                num_subfolders = folder.number_of_sub_folders

            for subfolder_index in range(num_subfolders):
                subfolder = folder.get_sub_folder(subfolder_index)
                if subfolder:
                    explore_folder(subfolder, level + 1, max_level)
        except Exception as e:
            print(f"{indent}  ❌ Error explorando subcarpetas: {e}")

def main():
    if len(sys.argv) < 2:
        print("Uso: python debug_pst.py <archivo_pst>")
        return

    pst_file_path = sys.argv[1]

    try:
        print(f"🔍 Explorando estructura de: {pst_file_path}")
        print("=" * 60)

        # Abrir archivo PST
        pst_file = pypff.file()
        pst_file.open(pst_file_path)

        # Obtener carpeta raíz
        root_folder = pst_file.get_root_folder()

        # Explorar estructura
        explore_folder(root_folder)

        print("=" * 60)
        print("✅ Exploración completada")

        pst_file.close()

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()