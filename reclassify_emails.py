#!/usr/bin/env python3
"""
Script para Re-clasificar Emails Existentes
Aplica los nuevos criterios de clasificación mejorados
"""

import json
import os
from pathlib import Path
from email_classifier import EmailClassifier
from datetime import datetime

def reclassify_all_emails():
    """Re-clasificar todos los emails con los criterios mejorados"""
    print("🔄 Iniciando re-clasificación de emails con criterios mejorados...")
    print("=" * 60)

    # Inicializar el clasificador
    classifier = EmailClassifier()

    # Cargar el archivo de clasificación existente
    classification_file = classifier.classification_dir / "classification_results.json"

    if not classification_file.exists():
        print("❌ Error: No se encontró el archivo de clasificación existente")
        return

    with open(classification_file, 'r', encoding='utf-8') as f:
        existing_data = json.load(f)

    print(f"📧 Emails a re-clasificar: {len(existing_data.get('emails', []))}")

    # Estadísticas antes de la re-clasificación
    stats_before = {
        'cotizacion': 0,
        'renovacion': 0,
        'endoso': 0,
        'sin_clasificar': 0
    }

    for email_info in existing_data.get('emails', []):
        classification_type = email_info['primary_classification']['type']
        if classification_type in stats_before:
            stats_before[classification_type] += 1

    print("\n📊 Clasificación ANTES:")
    for tipo, count in stats_before.items():
        print(f"   {tipo.capitalize()}: {count}")

    # Re-clasificar cada email
    reclassified_emails = []
    improved_count = 0

    for i, email_info in enumerate(existing_data.get('emails', [])):
        email_id = email_info['email_id']

        # Mostrar progreso
        if (i + 1) % 100 == 0:
            print(f"   Procesado: {i + 1}/{len(existing_data['emails'])}")

        # Re-clasificar con los nuevos criterios
        new_classification = classifier.classify_email(email_id)

        if 'error' not in new_classification:
            old_type = email_info['primary_classification']['type']
            new_type = new_classification['primary_classification']['type']

            # Verificar si hubo mejora en la clasificación
            if old_type == 'sin_clasificar' and new_type != 'sin_clasificar':
                improved_count += 1
                print(f"✅ Mejorado: {email_id} - {old_type} → {new_type} ({new_classification['primary_classification']['confidence']}%)")
            elif old_type != new_type:
                print(f"🔄 Cambio: {email_id} - {old_type} → {new_type}")

            reclassified_emails.append(new_classification)
        else:
            # Mantener clasificación anterior si hay error
            reclassified_emails.append(email_info)

    # Estadísticas después de la re-clasificación
    stats_after = {
        'cotizacion': 0,
        'renovacion': 0,
        'endoso': 0,
        'sin_clasificar': 0
    }

    for email_info in reclassified_emails:
        classification_type = email_info['primary_classification']['type']
        if classification_type in stats_after:
            stats_after[classification_type] += 1

    print(f"\n📊 Clasificación DESPUÉS:")
    for tipo, count in stats_after.items():
        print(f"   {tipo.capitalize()}: {count}")

    print(f"\n🎯 Emails mejorados: {improved_count}")

    # Crear nuevo archivo de resultados
    new_results = {
        'classification_summary': {
            'total_emails': len(reclassified_emails),
            'cotizacion': stats_after['cotizacion'],
            'renovacion': stats_after['renovacion'],
            'endoso': stats_after['endoso'],
            'sin_clasificar': stats_after['sin_clasificar'],
            'accuracy_improvement': f"{improved_count} emails mejorados",
            'reclassification_date': datetime.now().isoformat()
        },
        'classification_criteria': {
            'threshold_lowered': 'Umbral bajado de 50% a 30%',
            'improved_patterns': 'Patrones de endoso mejorados',
            'html_analysis': 'Análisis de contenido HTML implementado',
            'body_analysis': 'Análisis del cuerpo del email añadido'
        },
        'emails': reclassified_emails
    }

    # Guardar backup del archivo anterior
    backup_file = classifier.classification_dir / f"classification_results_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    print(f"\n💾 Guardando backup en: {backup_file}")

    with open(backup_file, 'w', encoding='utf-8') as f:
        json.dump(existing_data, f, indent=2, ensure_ascii=False)

    # Guardar nuevos resultados
    print(f"💾 Guardando resultados mejorados en: {classification_file}")

    with open(classification_file, 'w', encoding='utf-8') as f:
        json.dump(new_results, f, indent=2, ensure_ascii=False)

    print("\n✅ Re-clasificación completada!")
    print("=" * 60)

    # Mostrar resumen de mejoras
    print(f"""
🎉 RESUMEN DE MEJORAS:
   📈 Emails que mejoraron: {improved_count}
   📉 Sin clasificar antes: {stats_before['sin_clasificar']}
   📉 Sin clasificar después: {stats_after['sin_clasificar']}
   💪 Reducción sin clasificar: {stats_before['sin_clasificar'] - stats_after['sin_clasificar']}

🔧 MEJORAS IMPLEMENTADAS:
   • Umbral de confianza: 50% → 30%
   • Patrones de endoso ampliados (números OT, documentos, modificaciones)
   • Análisis del contenido HTML del email
   • Búsqueda en cuerpo del mensaje además del asunto
   • Mejor detección de códigos de agente y pólizas
    """)

if __name__ == "__main__":
    reclassify_all_emails()