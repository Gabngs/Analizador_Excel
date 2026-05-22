#!/bin/bash
# Script de compilación para Linux/macOS
# Compila Excel Analyzer en un ejecutable único

echo "🔨 Compilando Excel Analyzer v1.0.0 para $(uname -s)..."

# Verificar si PyInstaller está instalado
if ! command -v pyinstaller &> /dev/null
then
    echo "📦 Instalando PyInstaller..."
    pip install pyinstaller
fi

# Limpiar compilaciones anteriores
echo "🧹 Limpiando archivos anteriores..."
rm -rf build dist

# Compilar con PyInstaller
echo "⚙️  Compilando aplicación..."
pyinstaller excel_analyzer.spec

# Verificar si la compilación fue exitosa
if [ -f "dist/ExcelAnalyzer" ]; then
    echo "✅ Compilación exitosa!"
    echo "📁 Ejecutable creado en: dist/ExcelAnalyzer"
    
    # Hacer el ejecutable... ejecutable
    chmod +x dist/ExcelAnalyzer
    
    # Mostrar información del archivo
    ls -lh dist/ExcelAnalyzer
    
    echo ""
    echo "🚀 Para ejecutar:"
    echo "   ./dist/ExcelAnalyzer"
else
    echo "❌ Error en la compilación"
    exit 1
fi
