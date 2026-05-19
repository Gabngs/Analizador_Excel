#!/usr/bin/env python3
"""
Excel Analyzer CLI - Herramienta de línea de comandos
Uso:
  python3 cli.py analyze archivo.xlsx
  python3 cli.py compare archivo1.xlsx archivo2.xlsx
"""

import sys
import json
import argparse
from pathlib import Path
from excel_analyzer import ExcelAnalyzer

def print_header(text):
    print(f"\n{'='*60}")
    print(f"  {text}")
    print(f"{'='*60}\n")

def format_bytes(bytes_val):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_val < 1024:
            return f"{bytes_val:.2f} {unit}"
        bytes_val /= 1024
    return f"{bytes_val:.2f} TB"

def analyze_command(args):
    """Analiza un archivo Excel"""
    file_path = args.file
    
    if not Path(file_path).exists():
        print(f"❌ Error: Archivo no encontrado: {file_path}")
        return 1
    
    print_header(f"Análisis: {Path(file_path).name}")
    
    analyzer = ExcelAnalyzer()
    result = analyzer.analyze_single_file(file_path)
    
    if 'error' in result:
        print(f"❌ Error: {result['error']}")
        return 1
    
    print(f"📊 INFORMACIÓN GENERAL")
    print(f"  Archivo: {result['file_name']}")
    print(f"  Filas: {result['rows']:,}")
    print(f"  Columnas: {result['columns']}")
    print(f"  Memoria: {result['memory_usage']} KB")
    print(f"  Duplicados: {result['duplicates']}")
    
    print(f"\n📋 COLUMNAS ({result['columns']})")
    for col in result['column_names']:
        missing = result['missing_values'][col]
        dtype = result['dtypes'][col]
        missing_pct = (missing / result['rows'] * 100) if result['rows'] > 0 else 0
        status = "⚠️ " if missing > 0 else "✓"
        print(f"  {status} {col:<20} | Tipo: {dtype:<10} | Vacíos: {missing} ({missing_pct:.1f}%)")
    
    if result.get('numeric_stats'):
        print(f"\n📈 ESTADÍSTICAS NUMÉRICAS")
        for col, stats in result['numeric_stats'].items():
            print(f"  {col}")
            print(f"    Min: {stats['min']:,}")
            print(f"    Max: {stats['max']:,}")
            print(f"    Media: {stats['mean']:,.2f}")
    
    print()
    return 0

def compare_command(args):
    """Compara dos archivos Excel"""
    file1 = args.file1
    file2 = args.file2
    
    if not Path(file1).exists():
        print(f"❌ Error: Archivo no encontrado: {file1}")
        return 1
    if not Path(file2).exists():
        print(f"❌ Error: Archivo no encontrado: {file2}")
        return 1
    
    print_header(f"Comparación: {Path(file1).name} vs {Path(file2).name}")
    
    analyzer = ExcelAnalyzer()
    result = analyzer.compare_files(file1, file2)
    
    if 'error' in result:
        print(f"❌ Error: {result['error']}")
        return 1
    
    metrics = result['metrics']
    schema = result['schema_differences']
    data_diff = result['data_differences']
    
    print(f"📊 COMPARACIÓN DE TAMAÑO")
    print(f"  Archivo 1: {metrics['file1_rows']:,} filas × {metrics['file1_cols']} columnas")
    print(f"  Archivo 2: {metrics['file2_rows']:,} filas × {metrics['file2_cols']} columnas")
    print(f"  Diferencia de filas: {metrics['row_difference']:+,}")
    print(f"  Diferencia de columnas: {metrics['col_difference']:+,}")
    
    print(f"\n🔍 ESQUEMA")
    print(f"  Columnas comunes: {schema['common_count']}")
    print(f"  Columnas diferentes: {schema['total_different']}")
    
    if schema['only_in_file1']:
        print(f"\n  ❌ Solo en Archivo 1 ({len(schema['only_in_file1'])}):")
        for col in schema['only_in_file1']:
            print(f"     - {col}")
    
    if schema['only_in_file2']:
        print(f"\n  ❌ Solo en Archivo 2 ({len(schema['only_in_file2'])}):")
        for col in schema['only_in_file2']:
            print(f"     - {col}")
    
    if schema['type_changes']:
        print(f"\n  ⚠️ Cambios de Tipo ({len(schema['type_changes'])}):")
        for col, change in schema['type_changes'].items():
            print(f"     - {col}: {change['from']} → {change['to']}")
    
    print(f"\n📈 DIFERENCIAS EN DATOS")
    if data_diff.get('missing_in_file2'):
        print(f"  Filas faltantes en Archivo 2: {data_diff['missing_in_file2']}")
    if data_diff.get('extra_in_file2'):
        print(f"  Filas extra en Archivo 2: {data_diff['extra_in_file2']}")
    
    if data_diff['column_differences']:
        print(f"\n  Columnas con diferencias ({len(data_diff['column_differences'])}):")
        for col, count in sorted(data_diff['column_differences'].items(), 
                                key=lambda x: x[1], reverse=True):
            pct = (count / data_diff['rows_with_differences'] * 100) if data_diff['rows_with_differences'] > 0 else 0
            print(f"     - {col}: {count} celdas diferentes ({pct:.1f}%)")
    else:
        print(f"\n  ✅ Los datos son idénticos en columnas comunes")
    
    print()
    
    if args.json:
        print("\n📄 JSON:")
        print(json.dumps(result, indent=2, default=str))
    
    return 0

def main():
    parser = argparse.ArgumentParser(
        description='Excel Analyzer - Analiza y compara archivos Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Ejemplos:
  python3 cli.py analyze archivo.xlsx
  python3 cli.py compare file1.xlsx file2.xlsx --json
        '''
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Comando a ejecutar')
    
    # Comando analyze
    analyze_parser = subparsers.add_parser('analyze', help='Analiza un archivo Excel')
    analyze_parser.add_argument('file', help='Ruta del archivo Excel')
    analyze_parser.set_defaults(func=analyze_command)
    
    # Comando compare
    compare_parser = subparsers.add_parser('compare', help='Compara dos archivos Excel')
    compare_parser.add_argument('file1', help='Primer archivo Excel')
    compare_parser.add_argument('file2', help='Segundo archivo Excel')
    compare_parser.add_argument('--json', action='store_true', help='Mostrar resultado en JSON')
    compare_parser.set_defaults(func=compare_command)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    return args.func(args)

if __name__ == '__main__':
    sys.exit(main())
