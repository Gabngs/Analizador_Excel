# 🚀 Inicio Rápido - Excel Analyzer

## Instalación en 1 minuto

### Linux / Mac
```bash
chmod +x run.sh
./run.sh
```

### Windows
```bash
run.bat
```

### Manual
```bash
pip install -r requirements.txt
python excel_analyzer.py
```

## Acceso
- Abre: **http://localhost:5000**

---

## Uso por CLI (Terminal)

### Analizar un archivo
```bash
python3 cli.py analyze archivo.xlsx
```

Salida:
```
📊 INFORMACIÓN GENERAL
  Archivo: archivo.xlsx
  Filas: 1000
  Columnas: 5
  Memoria: 125 KB
  Duplicados: 3

📋 COLUMNAS (5)
  ✓ ID      | Tipo: int64  | Vacíos: 0
  ✓ Nombre  | Tipo: str    | Vacíos: 0
  ...
```

### Comparar dos archivos
```bash
python3 cli.py compare archivo1.xlsx archivo2.xlsx
```

O con JSON:
```bash
python3 cli.py compare archivo1.xlsx archivo2.xlsx --json
```

---

## Características Principales

✅ **Análisis Individual**
- Estadísticas completas (filas, columnas, tipos)
- Detección de valores faltantes
- Filas duplicadas
- Estadísticas numéricas (min, max, media)

✅ **Comparación de Archivos**
- Diferencias de esquema (campos faltantes, nuevos)
- Cambios de tipo de datos
- Diferencias en valores
- Métricas automáticas

✅ **Optimización**
- Bajo consumo de RAM
- Soporta archivos de 500MB+
- Caché de comparaciones
- Procesamiento eficiente

---

## Archivos Incluidos

| Archivo | Descripción |
|---------|------------|
| `excel_analyzer.py` | Aplicación web principal |
| `cli.py` | Herramienta línea de comandos |
| `utils.py` | Utilidades avanzadas |
| `config.py` | Configuración |
| `requirements.txt` | Dependencias Python |
| `run.sh` | Script de inicio (Linux/Mac) |
| `run.bat` | Script de inicio (Windows) |
| `generate_test_files.py` | Generador de datos de prueba |

---

## Ejemplos de Uso

### Generar archivos de prueba grandes
```bash
python3 generate_test_files.py 100000
```
Crea archivos de 100,000 filas cada uno

### Usar en Python como librería
```python
from excel_analyzer import ExcelAnalyzer

analyzer = ExcelAnalyzer()

# Análisis
result = analyzer.analyze_single_file('archivo.xlsx')
print(result)

# Comparación
comparison = analyzer.compare_files('file1.xlsx', 'file2.xlsx')
print(comparison)
```

---

## Troubleshooting

### ❌ "Python not found"
```bash
# Instala Python desde python.org o usa tu gestor de paquetes
```

### ❌ "Module not found"
```bash
pip install -r requirements.txt
```

### ❌ "Puerto 5000 en uso"
Edita `excel_analyzer.py` línea final:
```python
app.run(port=5001)  # Usa puerto diferente
```
Luego accede a: http://localhost:5001

### ❌ "Archivo muy grande"
El analyzer soporta hasta 500MB. Para archivos mayores:
```python
# Procesa en chunks
chunks = analyzer.read_excel_chunk('file.xlsx', chunk_size=50000)
```

---

## Notas de Rendimiento

- Archivos hasta 500MB: ✅ Rápido
- RAM mínima recomendada: 4GB
- Procesamiento en chunks: Activado por defecto
- Caché de SQLite: Acelera comparaciones repetidas

---

## Soporte

Para reportar problemas o sugerencias:
1. Verifica el archivo `analyzer.log` si está habilitado
2. Prueba con archivos de ejemplo más pequeños
3. Reinicia la aplicación

---

Hecho con ❤️ - Excel Analyzer v1.0
