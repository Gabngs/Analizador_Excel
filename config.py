"""
Configuración de la aplicación Excel Analyzer
Modifica estos valores según tus necesidades
"""

# Flask
FLASK_HOST = '0.0.0.0'
FLASK_PORT = 5000
FLASK_DEBUG = False
FLASK_THREADED = True

# Archivos
MAX_FILE_SIZE = 500 * 1024 * 1024  # 500MB en bytes
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
CHUNK_SIZE = 10000  # Filas por chunk para lectura en memoria

# Almacenamiento temporal
USE_TEMP_DIR = True  # Usar directorio temporal del SO
KEEP_UPLOADED_FILES = False  # Eliminar archivos después del procesamiento

# Base de datos
USE_SQLITE_CACHE = True  # Cachear comparaciones en SQLite
DB_RETENTION_DAYS = 7  # Días antes de limpiar cache antiguo

# Análisis
CALCULATE_STATISTICS = True  # Calcular estadísticas numéricas
DETECT_DUPLICATES = True  # Detectar filas duplicadas
CHECK_MISSING_VALUES = True  # Analizar valores faltantes

# UI
DISPLAY_MEMORY_USAGE = True
DISPLAY_DATA_TYPES = True
DISPLAY_NUMERIC_STATS = True
THEME_COLOR_PRIMARY = '#667eea'
THEME_COLOR_SECONDARY = '#764ba2'

# Logging
LOG_ENABLED = True
LOG_LEVEL = 'INFO'
LOG_FILE = 'analyzer.log'

# Performance
MAX_ROWS_FOR_FULL_COMPARISON = 100000
ENABLE_PROGRESS_BAR = True
TIMEOUT_SECONDS = 300

# Seguridad
SANITIZE_FILENAMES = True
VALIDATE_FILE_HEADERS = True
