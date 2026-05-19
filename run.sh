#!/bin/bash

echo "=== Instalador Excel Analyzer ==="
echo ""

if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 no está instalado"
    exit 1
fi

echo "✓ Python encontrado: $(python3 --version)"
echo ""

echo "📦 Instalando dependencias..."
pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "✓ Dependencias instaladas correctamente"
    echo ""
    echo "🚀 Iniciando aplicación..."
    echo "📱 Abre en tu navegador: http://localhost:5000"
    echo ""
    python3 excel_analyzer.py
else
    echo "❌ Error al instalar dependencias"
    exit 1
fi
