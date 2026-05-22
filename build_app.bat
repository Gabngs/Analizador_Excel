@echo off
REM Script de compilación para Windows
REM Compila Excel Analyzer en un ejecutable único

echo 🔨 Compilando Excel Analyzer v1.0.0 para Windows...

REM Verificar si PyInstaller está instalado
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 📦 Instalando PyInstaller...
    pip install pyinstaller
)

REM Limpiar compilaciones anteriores
echo 🧹 Limpiando archivos anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Compilar con PyInstaller
echo ⚙️  Compilando aplicación...
pyinstaller excel_analyzer.spec

REM Verificar si la compilación fue exitosa
if exist "dist\ExcelAnalyzer.exe" (
    echo ✅ Compilación exitosa!
    echo 📁 Ejecutable creado en: dist\ExcelAnalyzer.exe
    dir dist\ExcelAnalyzer.exe
    echo.
    echo 🚀 Para ejecutar:
    echo    dist\ExcelAnalyzer.exe
) else (
    echo ❌ Error en la compilación
    exit /b 1
)
