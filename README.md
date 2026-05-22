# Excel Analyzer - Comparador y Analizador de Archivos Excel

Herramienta web para analizar y comparar archivos Excel con interfaz visual, optimizada para bajo consumo de RAM.

## 📥 Descargas Rápidas

### Ejecutables (No requiere Python instalado)

Descarga el ejecutable para tu sistema operativo desde [Releases](https://github.com/TU_USUARIO/TU_REPOSITORIO/releases/latest):

- **Windows**: `ExcelAnalyzer-windows-x64.exe`
- **Linux**: `ExcelAnalyzer-linux-x64`
- **macOS**: `ExcelAnalyzer-macos-x64`

#### Instrucciones:

1. Descarga el archivo para tu sistema
2. **Linux/macOS**: Dar permisos de ejecución: `chmod +x ExcelAnalyzer-*`
3. Ejecuta el archivo
4. Abre tu navegador en `http://localhost:5000`

## Características

- ✅ Análisis individual de archivos Excel
- ✅ Comparación automática de 2 archivos
- ✅ Detección de diferencias en campos (esquema)
- ✅ Detección de diferencias en datos
- ✅ Estadísticas automáticas
- ✅ Interfaz web responsiva
- ✅ Optimizado para archivos grandes (bajo consumo RAM)
- ✅ Soporta: XLSX, XLS, CSV

## Requisitos

- Python 3.7+
- pip (administrador de paquetes de Python)

## Instalación Rápida

### En Linux/Mac:

```bash
chmod +x run.sh
./run.sh
```

### En Windows:

```bash
run.bat
```

O manual:

```bash
pip install -r requirements.txt
python excel_analyzer.py
```

## Uso

1. Abre tu navegador en `http://localhost:5000`
2. Arrastra archivos Excel o selecciónalos
3. Elige analizar o comparar

## Funcionalidades

### Análisis Individual

- Número de filas y columnas
- Tipos de datos por columna
- Valores faltantes
- Filas duplicadas
- Uso de memoria
- Estadísticas numéricas (min, máx, media)

### Comparación de 2 Archivos

- Columnas solo en archivo 1
- Columnas solo en archivo 2
- Cambios de tipo de datos
- Diferencias en valores
- Métricas de diferencia

## Rendimiento

Optimizado para:

- Archivos de hasta 500MB
- Lectura en chunks para bajo consumo RAM
- Caché de comparaciones en SQLite
- Procesamiento eficiente sin cargar todo en memoria

## Archivos Generados

- `excel_analyzer.py` - Aplicación principal
- `requirements.txt` - Dependencias Python
- `run.sh` - Script de inicio (Linux/Mac)
- `run.bat` - Script de inicio (Windows)

## Solución de Problemas

Si no inicia:

1. Verifica que Python 3 esté instalado: `python --version`
2. Instala dependencias manualmente: `pip install -r requirements.txt`
3. Ejecuta directamente: `python excel_analyzer.py`

Si hay error de puerto:

- Cambia el puerto en excel_analyzer.py línea final: `port=5001`
- Luego accede a: `http://localhost:5001`

## 🔨 Para Desarrolladores

### Compilar Ejecutables Localmente

```bash
# Instalar dependencias de desarrollo
pip install -r requirements-dev.txt

# Windows
build_app.bat

# Linux/macOS
chmod +x build_app.sh
./build_app.sh
```

El ejecutable estará en la carpeta `dist/`

### Crear un Release

Ver [RELEASE_GUIDE.md](RELEASE_GUIDE.md) para instrucciones detalladas sobre cómo crear y publicar versiones.

## Licencia

Uso libre para propósitos comerciales y personales.
