#!/usr/bin/env python3
"""
Script de debug avanzado para inspeccionar métodos disponibles en pypff
"""

import pypff
import sys

def inspect_object(obj, name):
    """Inspecciona los métodos y atributos disponibles de un objeto"""
    print(f"\n🔬 Inspeccionando {name}:")
    methods = [method for method in dir(obj) if not method.startswith('_')]

    # Buscar métodos relacionados con mensajes
    message_methods = [m for m in methods if 'message' in m.lower()]
    if message_methods:
        print(f"  📧 Métodos de mensajes: {message_methods}")

    # Buscar métodos relacionados con items
    item_methods = [m for m in methods if 'item' in m.lower()]
    if item_methods:
        print(f"  📦 Métodos de items: {item_methods}")

    # Buscar métodos relacionados con count/number
    count_methods = [m for m in methods if any(word in m.lower() for word in ['count', 'number', 'size'])]
    if count_methods:
        print(f"  🔢 Métodos de conteo: {count_methods}")

def explore_folder_advanced(folder, level=0, max_level=2):
    """Explora carpetas con inspección avanzada"""
    indent = "  " * level

    # Obtener nombre de la carpeta
    folder_name = "ROOT"
    if hasattr(folder, 'name') and folder.name:
        folder_name = folder.name

    print(f"{indent}📁 {folder_name}")

    # Inspeccionar carpetas importantes
    if folder_name in ['ASIGNADOS', 'PRODUCCION', 'DEPURADOS'] or level == 0:
        inspect_object(folder, f"Carpeta {folder_name}")

    # Probar diferentes métodos para obtener mensajes
    message_count = 0
    methods_tried = []

    # Método 1: get_number_of_messages
    try:
        if hasattr(folder, 'get_number_of_messages'):
            message_count = folder.get_number_of_messages()
            methods_tried.append(f"get_number_of_messages() = {message_count}")
    except Exception as e:
        methods_tried.append(f"get_number_of_messages() ERROR: {e}")

    # Método 2: number_of_messages
    try:
        if hasattr(folder, 'number_of_messages'):
            message_count = folder.number_of_messages
            methods_tried.append(f"number_of_messages = {message_count}")
    except Exception as e:
        methods_tried.append(f"number_of_messages ERROR: {e}")

    # Método 3: get_number_of_items
    try:
        if hasattr(folder, 'get_number_of_items'):
            item_count = folder.get_number_of_items()
            methods_tried.append(f"get_number_of_items() = {item_count}")
    except Exception as e:
        methods_tried.append(f"get_number_of_items() ERROR: {e}")

    # Método 4: Iteración directa
    try:
        direct_count = 0
        while True:
            try:
                msg = folder.get_message(direct_count)
                if msg:
                    direct_count += 1
                else:
                    break
            except:
                break
        if direct_count > 0:
            methods_tried.append(f"iteración_directa = {direct_count}")
    except Exception as e:
        methods_tried.append(f"iteración_directa ERROR: {e}")

    # Mostrar resultados
    for method_result in methods_tried:
        print(f"{indent}  🧪 {method_result}")

    if message_count > 0:
        print(f"{indent}  ✅ {message_count} mensajes encontrados!")

        # Intentar leer el primer mensaje para verificar
        try:
            first_message = folder.get_message(0)
            if first_message:
                subject = "Sin asunto"
                if hasattr(first_message, 'subject') and first_message.subject:
                    subject = first_message.subject[:50] + "..." if len(first_message.subject) > 50 else first_message.subject
                print(f"{indent}    📬 Primer mensaje: {subject}")
        except Exception as e:
            print(f"{indent}    ❌ Error leyendo primer mensaje: {e}")

    # Explorar subcarpetas
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
                    explore_folder_advanced(subfolder, level + 1, max_level)
        except Exception as e:
            print(f"{indent}  ❌ Error explorando subcarpetas: {e}")

def main():
    if len(sys.argv) < 2:
        print("Uso: python debug_pst_advanced.py <archivo_pst>")
        return

    pst_file_path = sys.argv[1]

    try:
        print(f"🔬 Análisis avanzado de: {pst_file_path}")
        print("=" * 80)

        # Abrir archivo PST
        pst_file = pypff.file()
        pst_file.open(pst_file_path)

        # Obtener carpeta raíz
        root_folder = pst_file.get_root_folder()

        # Explorar estructura
        explore_folder_advanced(root_folder)

        print("=" * 80)
        print("✅ Análisis completado")

        pst_file.close()

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()