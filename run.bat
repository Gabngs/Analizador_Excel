@echo off
setlocal enabledelayedexpansion

echo === Instalador Excel Analyzer ===
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python 3 no está instalado
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo ✓ Python encontrado: %PYTHON_VERSION%
echo.

echo 📦 Instalando dependencias...
pip install -r requirements.txt

if errorlevel 1 (
    echo ❌ Error al instalar dependencias
    pause
    exit /b 1
)

echo.
echo ✓ Dependencias instaladas correctamente
echo.
echo 🚀 Iniciando aplicación...
echo 📱 Abre en tu navegador: http://localhost:5000
echo.
python excel_analyzer.py

pause
