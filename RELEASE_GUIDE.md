# 🚀 Guía para Crear y Publicar Releases

Esta guía te explica paso a paso cómo compilar Excel Analyzer y publicar la versión 1.0.0 (o cualquier versión) para que sea descargable en Windows, Linux y macOS.

## 📋 Requisitos Previos

1. **Python 3.7+** instalado
2. **Git** instalado
3. Una cuenta de **GitHub**
4. El proyecto subido a un repositorio de GitHub

## 🛠️ Opción 1: Compilación Automática (Recomendado)

Esta opción usa GitHub Actions para compilar automáticamente en las 3 plataformas.

### Paso 1: Subir el código a GitHub

```bash
# Si no has inicializado git todavía:
git init
git add .
git commit -m "Initial commit - Excel Analyzer v1.0.0"

# Conectar con tu repositorio de GitHub
git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
git branch -M main
git push -u origin main
```

### Paso 2: Crear un tag de versión

```bash
# Crear el tag v1.0.0
git tag -a v1.0.0 -m "Release version 1.0.0"

# Subir el tag a GitHub
git push origin v1.0.0
```

### Paso 3: GitHub Actions compila automáticamente

- Ve a tu repositorio en GitHub
- Navega a la pestaña **"Actions"**
- Verás que se está ejecutando el workflow "Build and Release"
- Espera unos 10-15 minutos mientras compila en Windows, Linux y macOS

### Paso 4: Descargar los ejecutables

- Una vez completado, ve a la pestaña **"Releases"** en GitHub
- Verás el release **v1.0.0** con 3 archivos:
  - `ExcelAnalyzer-windows-x64.exe` (Windows)
  - `ExcelAnalyzer-linux-x64` (Linux)
  - `ExcelAnalyzer-macos-x64` (macOS)

## 🔧 Opción 2: Compilación Local

Si prefieres compilar manualmente en tu computadora:

### En Windows:

```bash
# Instalar PyInstaller
pip install pyinstaller

# Ejecutar el script de compilación
build_app.bat

# El ejecutable estará en: dist\ExcelAnalyzer.exe
```

### En Linux/macOS:

```bash
# Instalar PyInstaller
pip install pyinstaller

# Dar permisos al script
chmod +x build_app.sh

# Ejecutar el script de compilación
./build_app.sh

# El ejecutable estará en: dist/ExcelAnalyzer
```

### Subir manualmente el ejecutable

Después de compilar localmente:

1. Ve a tu repositorio en GitHub
2. Click en **"Releases"** → **"Create a new release"**
3. En "Tag version" escribe: `v1.0.0`
4. En "Release title" escribe: `Excel Analyzer v1.0.0`
5. Arrastra y suelta el ejecutable compilado
6. Click en **"Publish release"**

## 📦 Crear Nuevas Versiones

Para crear versiones futuras (v1.1.0, v2.0.0, etc.):

### 1. Actualizar el número de versión

Edita el archivo `version.py`:

```python
__version__ = "1.1.0"  # Cambia esto
```

### 2. Commit y crear nuevo tag

```bash
git add .
git commit -m "Update to version 1.1.0"
git tag -a v1.1.0 -m "Release version 1.1.0"
git push origin main
git push origin v1.1.0
```

### 3. GitHub Actions compila automáticamente

El proceso se repite automáticamente para la nueva versión.

## 🎯 Versionado Semántico

Sigue este formato para versiones: `MAJOR.MINOR.PATCH`

- **MAJOR** (1.0.0 → 2.0.0): Cambios incompatibles
- **MINOR** (1.0.0 → 1.1.0): Nuevas funciones compatibles
- **PATCH** (1.0.0 → 1.0.1): Correcciones de bugs

Ejemplos:

- `v1.0.0` - Primera versión estable
- `v1.1.0` - Agregaste análisis de gráficos
- `v1.0.1` - Corregiste un bug de memoria
- `v2.0.0` - Cambiaste la interfaz completamente

## ✅ Verificar el Release

Después de publicar, verifica:

1. ✅ Los 3 archivos están disponibles para descargar
2. ✅ Cada ejecutable funciona en su plataforma
3. ✅ El README tiene instrucciones claras
4. ✅ El tag de versión es correcto

## 🔍 Troubleshooting

### Error: "GITHUB_TOKEN permissions"

Si falla el release automático, ve a:

- Repositorio → Settings → Actions → General
- "Workflow permissions" → Marca "Read and write permissions"

### Error en compilación

- Verifica que `requirements.txt` tenga todas las dependencias
- Asegúrate de que el código funciona localmente primero
- Revisa los logs en GitHub Actions

## 📊 Estructura de Archivos Generados

```
Releases/
├── v1.0.0/
│   ├── ExcelAnalyzer-windows-x64.exe    (~50-80 MB)
│   ├── ExcelAnalyzer-linux-x64          (~50-80 MB)
│   └── ExcelAnalyzer-macos-x64          (~50-80 MB)
```

## 🌐 Compartir con Usuarios

Comparte el link directo:

```
https://github.com/TU_USUARIO/TU_REPOSITORIO/releases/latest
```

Los usuarios podrán:

1. Ver todas las versiones disponibles
2. Descargar el ejecutable para su sistema
3. Ejecutarlo sin instalar Python

## 📝 Notas Importantes

- **Tamaño**: Los ejecutables pesan ~50-80 MB porque incluyen Python y todas las librerías
- **Antivirus**: Windows puede mostrar advertencia en ejecutables no firmados (es normal)
- **Permisos**: En Linux/macOS los usuarios deben ejecutar `chmod +x` antes de usar
- **Puerto**: La app usa el puerto 5000, asegúrate de que esté libre

## 🎉 ¡Listo!

Ahora tienes un sistema de releases automático que:

- ✅ Compila en 3 sistemas operativos
- ✅ Crea releases automáticamente
- ✅ Permite descargas fáciles
- ✅ Mantiene historial de versiones
