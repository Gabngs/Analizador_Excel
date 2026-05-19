#!/usr/bin/env python3
"""
Generador de archivos Excel grandes para testing
Uso: python3 generate_test_files.py [num_rows]
"""

import pandas as pd
import numpy as np
import sys

def generate_test_file(filename, num_rows=100000):
    """Genera archivo Excel con datos aleatorios"""
    print(f"Generando {filename} con {num_rows} filas...")
    
    data = {
        'ID': np.arange(1, num_rows + 1),
        'Nombre': [f'Usuario_{i}' for i in range(num_rows)],
        'Email': [f'user_{i}@example.com' for i in range(num_rows)],
        'Edad': np.random.randint(18, 80, num_rows),
        'Salario': np.random.randint(30000, 150000, num_rows),
        'Departamento': np.random.choice(['IT', 'HR', 'Sales', 'Finance', 'Marketing'], num_rows),
        'Fecha_Ingreso': pd.date_range('2020-01-01', periods=num_rows, freq='H'),
        'Activo': np.random.choice([True, False], num_rows),
        'Score': np.random.random(num_rows),
        'Ventas': np.random.randint(0, 500000, num_rows)
    }
    
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False, engine='openpyxl')
    
    file_size_mb = df.memory_usage(deep=True).sum() / (1024 * 1024)
    print(f"✓ {filename} creado ({file_size_mb:.2f} MB en memoria)")

if __name__ == '__main__':
    rows = int(sys.argv[1]) if len(sys.argv) > 1 else 50000
    
    generate_test_file('large_file_1.xlsx', rows)
    
    # Crear variación del archivo
    print("\nGenerando segunda variante...")
    data = {
        'ID': np.arange(1, rows + 101),
        'Nombre': [f'Usuario_{i}' for i in range(rows + 100)],
        'Email': [f'user_{i}@example.com' for i in range(rows + 100)],
        'Edad': np.random.randint(18, 80, rows + 100),
        'Salario': np.random.randint(30000, 150000, rows + 100),
        'Departamento': np.random.choice(['IT', 'HR', 'Sales', 'Finance', 'Marketing'], rows + 100),
        'Fecha_Ingreso': pd.date_range('2020-01-01', periods=rows + 100, freq='H'),
        'Activo': np.random.choice([True, False], rows + 100),
        'Score': np.random.random(rows + 100),
        'Ventas': np.random.randint(0, 500000, rows + 100),
        'Region': np.random.choice(['North', 'South', 'East', 'West'], rows + 100)
    }
    
    df = pd.DataFrame(data)
    df.to_excel('large_file_2.xlsx', index=False, engine='openpyxl')
    
    print("✓ Archivos de prueba generados exitosamente")
