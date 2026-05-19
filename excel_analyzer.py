import os
import json
import sqlite3
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Tuple
import pandas as pd
import numpy as np
from flask import Flask, render_template_string, request, jsonify
from werkzeug.utils import secure_filename
import hashlib

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

class ExcelAnalyzer:
    def __init__(self):
        self.db_path = os.path.join(tempfile.gettempdir(), 'analyzer.db')
        self.init_db()
        
    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute('''CREATE TABLE IF NOT EXISTS comparisons
                       (id TEXT PRIMARY KEY, data TEXT)''')
        conn.commit()
        conn.close()
    
    def read_excel_chunk(self, file_path: str, chunk_size: int = 10000):
        """Lee archivo Excel en chunks para optimizar RAM"""
        chunks = []
        try:
            for chunk in pd.read_excel(file_path, sheet_name=None, 
                                      chunksize=chunk_size):
                chunks.append(chunk)
            return chunks if chunks else None
        except:
            return None
    
    def get_sheet_info(self, file_path: str) -> Dict[str, Any]:
        """Extrae info de hojas sin cargar todo en RAM"""
        try:
            xls = pd.ExcelFile(file_path)
            sheets = {}
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
                sheets[sheet_name] = {
                    'columns': df.columns.tolist(),
                    'dtypes': df.dtypes.astype(str).to_dict(),
                    'shape': pd.read_excel(file_path, sheet_name=sheet_name).shape
                }
            return sheets
        except Exception as e:
            return {'error': str(e)}
    
    def analyze_single_file(self, file_path: str) -> Dict[str, Any]:
        """Análisis de un archivo individual"""
        try:
            df = pd.read_excel(file_path)
            
            metrics = {
                'file_name': Path(file_path).name,
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': df.columns.tolist(),
                'dtypes': {col: str(dtype) for col, dtype in df.dtypes.items()},
                'missing_values': {col: int(df[col].isna().sum()) for col in df.columns},
                'duplicates': int(df.duplicated().sum()),
                'memory_usage': int(df.memory_usage(deep=True).sum() / 1024),
                'shape_info': f"{len(df)} filas × {len(df.columns)} columnas"
            }
            
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                metrics['numeric_stats'] = {
                    col: {
                        'mean': float(df[col].mean()) if pd.notna(df[col].mean()) else None,
                        'min': float(df[col].min()) if pd.notna(df[col].min()) else None,
                        'max': float(df[col].max()) if pd.notna(df[col].max()) else None,
                    }
                    for col in numeric_cols
                }
            
            return metrics
        except Exception as e:
            return {'error': str(e)}
    
    def compare_files(self, file1_path: str, file2_path: str, debug_mode: bool = False) -> Dict[str, Any]:
        """Compara dos archivos Excel"""
        try:
            df1 = pd.read_excel(file1_path)
            df2 = pd.read_excel(file2_path)
            
            comparison = {
                'file1': Path(file1_path).name,
                'file2': Path(file2_path).name,
                'timestamp': pd.Timestamp.now().isoformat(),
                'schema_differences': self._compare_schemas(df1, df2),
                'data_differences': self._compare_data(df1, df2),
                'metrics': {
                    'file1_rows': len(df1),
                    'file2_rows': len(df2),
                    'file1_cols': len(df1.columns),
                    'file2_cols': len(df2.columns),
                    'row_difference': len(df2) - len(df1),
                    'col_difference': len(df2.columns) - len(df1.columns),
                }
            }
            
            # Agregar ejemplos de diferencias para depuración
            if debug_mode:
                comparison['debug_samples'] = self._get_difference_samples(df1, df2)
            
            comp_id = hashlib.md5(
                f"{Path(file1_path).name}{Path(file2_path).name}".encode()
            ).hexdigest()
            
            conn = sqlite3.connect(self.db_path)
            conn.execute('INSERT OR REPLACE INTO comparisons VALUES (?, ?)',
                        (comp_id, json.dumps(comparison)))
            conn.commit()
            conn.close()
            
            return comparison
        except Exception as e:
            return {'error': str(e)}
    
    def _compare_schemas(self, df1: pd.DataFrame, df2: pd.DataFrame) -> Dict[str, Any]:
        """Compara esquemas (columnas y tipos)"""
        cols1 = set(df1.columns)
        cols2 = set(df2.columns)
        
        return {
            'only_in_file1': list(cols1 - cols2),
            'only_in_file2': list(cols2 - cols1),
            'common_columns': list(cols1 & cols2),
            'type_changes': self._get_type_changes(df1, df2, cols1 & cols2),
            'common_count': len(cols1 & cols2),
            'total_different': len(cols1 ^ cols2)
        }
    
    def _get_type_changes(self, df1, df2, common_cols):
        """Detecta cambios de tipo en columnas comunes"""
        changes = {}
        for col in common_cols:
            type1 = str(df1[col].dtype)
            type2 = str(df2[col].dtype)
            if type1 != type2:
                changes[col] = {'from': type1, 'to': type2}
        return changes
    
    def _normalize_value(self, val):
        """Normaliza un valor para comparación"""
        # Si es NaN/None, retornar un valor especial
        if pd.isna(val):
            return None
        
        # Si es string, limpiar espacios y convertir a minúsculas
        if isinstance(val, str):
            return val.strip()
        
        # Si es número, convertir a float para comparación consistente
        if isinstance(val, (int, float, np.integer, np.floating)):
            return float(val)
        
        # Si es fecha/timestamp, convertir a string normalizado
        if isinstance(val, (pd.Timestamp, np.datetime64)):
            return pd.Timestamp(val).isoformat()
        
        return val
    
    def _get_difference_samples(self, df1: pd.DataFrame, df2: pd.DataFrame, max_samples: int = 5) -> Dict[str, List]:
        """Obtiene ejemplos de diferencias para diagnóstico"""
        common_cols = set(df1.columns) & set(df2.columns)
        samples = {}
        
        for col in list(common_cols)[:10]:  # Limitar a 10 columnas para no saturar
            col_samples = []
            min_len = min(len(df1), len(df2))
            
            for i in range(min(min_len, 100)):  # Revisar primeras 100 filas
                val1 = df1[col].iloc[i]
                val2 = df2[col].iloc[i]
                
                # Verificar si son diferentes
                are_different = False
                if pd.isna(val1) and pd.isna(val2):
                    continue
                elif pd.isna(val1) or pd.isna(val2):
                    are_different = True
                else:
                    norm1 = self._normalize_value(val1)
                    norm2 = self._normalize_value(val2)
                    are_different = norm1 != norm2
                
                if are_different and len(col_samples) < max_samples:
                    col_samples.append({
                        'row': i,
                        'file1_value': str(val1),
                        'file2_value': str(val2),
                        'file1_type': str(type(val1).__name__),
                        'file2_type': str(type(val2).__name__),
                        'file1_repr': repr(val1),
                        'file2_repr': repr(val2)
                    })
            
            if col_samples:
                samples[col] = col_samples
        
        return samples
    
    def _compare_data(self, df1: pd.DataFrame, df2: pd.DataFrame) -> Dict[str, Any]:
        """Compara datos entre archivos"""
        common_cols = set(df1.columns) & set(df2.columns)
        
        if len(common_cols) == 0:
            return {'error': 'No hay columnas comunes'}
        
        df1_common = df1[list(common_cols)].reset_index(drop=True)
        df2_common = df2[list(common_cols)].reset_index(drop=True)
        
        differences = {
            'missing_in_file2': int(len(df1_common) - len(df2_common)) if len(df1_common) > len(df2_common) else 0,
            'extra_in_file2': int(len(df2_common) - len(df1_common)) if len(df2_common) > len(df1_common) else 0,
            'rows_with_differences': 0,
            'column_differences': {}
        }
        
        min_len = min(len(df1_common), len(df2_common))
        
        for col in common_cols:
            if col in df1_common.columns and col in df2_common.columns:
                diff_count = 0
                if min_len > 0:
                    col1 = df1_common[col].iloc[:min_len]
                    col2 = df2_common[col].iloc[:min_len]
                    
                    # Normalizar valores antes de comparar
                    col1_normalized = col1.apply(self._normalize_value)
                    col2_normalized = col2.apply(self._normalize_value)
                    
                    # Comparar valores normalizados — excluir NaN vs NaN explícitamente
                    # (pandas reconvierte None a nan en la Serie, y nan != nan es True en IEEE 754)
                    both_nan = col1_normalized.isna() & col2_normalized.isna()
                    diff_mask = (col1_normalized != col2_normalized) & ~both_nan

                    # Para valores numéricos, usar tolerancia para flotantes
                    if col1.dtype in ['float64', 'float32'] and col2.dtype in ['float64', 'float32']:
                        # Comparar con tolerancia para errores de precisión flotante
                        numeric_diff = ~np.isclose(
                            col1.fillna(-999999), 
                            col2.fillna(-999999), 
                            rtol=1e-9, 
                            atol=1e-9, 
                            equal_nan=True
                        )
                        diff_mask = numeric_diff
                    
                    diff_count = int(diff_mask.sum())
                
                if diff_count > 0:
                    differences['column_differences'][col] = int(diff_count)
        
        differences['rows_with_differences'] = min_len
        
        return differences

analyzer = ExcelAnalyzer()

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Analyzer</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .content {
            padding: 40px;
        }
        .section {
            margin-bottom: 40px;
        }
        .section h2 {
            font-size: 1.5em;
            color: #333;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #667eea;
        }
        .upload-area {
            border: 2px dashed #667eea;
            border-radius: 8px;
            padding: 30px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
            background: #f8f9ff;
        }
        .upload-area:hover {
            border-color: #764ba2;
            background: #f0f2ff;
        }
        .upload-area.dragover {
            border-color: #764ba2;
            background: #e8ebff;
        }
        input[type="file"] {
            display: none;
        }
        .file-input-label {
            cursor: pointer;
            color: #667eea;
            font-weight: 600;
            text-decoration: underline;
        }
        .button {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 30px;
            border-radius: 6px;
            border: none;
            cursor: pointer;
            font-size: 1em;
            transition: transform 0.2s, box-shadow 0.2s;
            margin: 10px 5px 10px 0;
        }
        .button:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }
        .button:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }
        .file-list {
            list-style: none;
            margin: 20px 0;
        }
        .file-item {
            background: #f8f9ff;
            padding: 12px;
            margin: 8px 0;
            border-radius: 6px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .file-item span { color: #333; font-weight: 500; }
        .file-remove {
            background: #ff6b6b;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.9em;
        }
        .file-remove:hover { background: #ff5252; }
        .results {
            background: #f8f9ff;
            border-radius: 8px;
            padding: 25px;
            margin-top: 20px;
        }
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        .metric-card {
            background: white;
            border-left: 4px solid #667eea;
            padding: 15px;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .metric-card h3 {
            font-size: 0.9em;
            color: #666;
            margin-bottom: 8px;
            text-transform: uppercase;
        }
        .metric-card .value {
            font-size: 1.8em;
            color: #667eea;
            font-weight: bold;
        }
        .differences {
            background: white;
            border-radius: 6px;
            padding: 15px;
            margin: 15px 0;
            border-left: 4px solid #ff6b6b;
        }
        .differences h4 {
            color: #ff6b6b;
            margin-bottom: 10px;
        }
        .schema-diff {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin: 15px 0;
        }
        .schema-item {
            background: white;
            padding: 12px;
            border-radius: 6px;
            border: 1px solid #e0e0e0;
        }
        .schema-item h5 {
            color: #333;
            margin-bottom: 8px;
            font-size: 0.9em;
            text-transform: uppercase;
            color: #666;
        }
        .schema-item ul {
            list-style-position: inside;
            font-size: 0.9em;
            color: #555;
        }
        .schema-item li {
            padding: 4px 0;
            word-break: break-word;
        }
        .loading {
            display: none;
            text-align: center;
            padding: 30px;
        }
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .error {
            background: #ffe0e0;
            color: #c33;
            padding: 15px;
            border-radius: 6px;
            margin: 15px 0;
            border-left: 4px solid #ff6b6b;
        }
        .success {
            background: #e0ffe0;
            color: #3c3;
            padding: 15px;
            border-radius: 6px;
            margin: 15px 0;
            border-left: 4px solid #6bff6b;
        }
        .tab-buttons {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            border-bottom: 2px solid #e0e0e0;
        }
        .tab-btn {
            background: none;
            border: none;
            padding: 12px 20px;
            cursor: pointer;
            color: #666;
            font-weight: 500;
            border-bottom: 3px solid transparent;
            transition: all 0.3s;
        }
        .tab-btn.active {
            color: #667eea;
            border-bottom-color: #667eea;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
        }
        .comparison-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin: 20px 0;
        }
        .file-col {
            background: white;
            padding: 15px;
            border-radius: 6px;
            border: 1px solid #e0e0e0;
        }
        .file-col h4 {
            color: #667eea;
            margin-bottom: 15px;
            font-size: 1.1em;
        }
        @media (max-width: 768px) {
            .comparison-row { grid-template-columns: 1fr; }
            .metrics-grid { grid-template-columns: 1fr; }
            .header h1 { font-size: 1.8em; }
            .content { padding: 20px; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Excel Analyzer</h1>
            <p>Compara y analiza archivos Excel con precisión</p>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>Cargar Archivos</h2>
                <div class="upload-area" id="uploadArea" ondrop="handleDrop(event)" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)">
                    <p>Arrastra archivos aquí o <label class="file-input-label">selecciona desde tu PC</label></p>
                    <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" multiple onchange="handleFileSelect(event)">
                </div>
                
                <div>
                    <ul class="file-list" id="fileList"></ul>
                </div>
                
                <div style="margin-top: 20px;">
                    <button class="button" onclick="analyzeSingle()" id="singleBtn" disabled>Analizar Archivo</button>
                    <button class="button" onclick="compareFiles()" id="compareBtn" disabled>Comparar 2 Archivos</button>
                    <button class="button" onclick="clearFiles()" style="background: #999;">Limpiar</button>
                </div>
            </div>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>Procesando archivos...</p>
            </div>
            
            <div id="resultsContainer"></div>
        </div>
    </div>
    
    <script>
        let selectedFiles = [];
        const uploadArea = document.getElementById('uploadArea');
        
        function handleDragOver(e) {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        }
        
        function handleDragLeave(e) {
            uploadArea.classList.remove('dragover');
        }
        
        function handleDrop(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = Array.from(e.dataTransfer.files);
            addFiles(files);
        }
        
        function handleFileSelect(e) {
            const files = Array.from(e.target.files);
            addFiles(files);
        }
        
        function addFiles(files) {
            selectedFiles = [...selectedFiles, ...files].slice(0, 2);
            if (selectedFiles.length > 2) selectedFiles = selectedFiles.slice(0, 2);
            updateFileList();
        }
        
        function updateFileList() {
            const list = document.getElementById('fileList');
            list.innerHTML = '';
            selectedFiles.forEach((file, idx) => {
                const li = document.createElement('li');
                li.className = 'file-item';
                li.innerHTML = `
                    <span>${file.name} (${(file.size / 1024).toFixed(2)} KB)</span>
                    <button class="file-remove" onclick="removeFile(${idx})">Eliminar</button>
                `;
                list.appendChild(li);
            });
            
            document.getElementById('singleBtn').disabled = selectedFiles.length === 0;
            document.getElementById('compareBtn').disabled = selectedFiles.length !== 2;
        }
        
        function removeFile(idx) {
            selectedFiles.splice(idx, 1);
            updateFileList();
        }
        
        function clearFiles() {
            selectedFiles = [];
            updateFileList();
            document.getElementById('resultsContainer').innerHTML = '';
        }
        
        function showLoading() {
            document.getElementById('loading').style.display = 'block';
        }
        
        function hideLoading() {
            document.getElementById('loading').style.display = 'none';
        }
        
        function analyzeSingle() {
            if (selectedFiles.length === 0) {
                showError('Selecciona al menos un archivo');
                return;
            }
            
            showLoading();
            const formData = new FormData();
            formData.append('file', selectedFiles[0]);
            
            fetch('/analyze-single', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    hideLoading();
                    if (data.error) {
                        showError(data.error);
                    } else {
                        displaySingleAnalysis(data);
                    }
                })
                .catch(e => {
                    hideLoading();
                    showError('Error: ' + e.message);
                });
        }
        
        function compareFiles() {
            if (selectedFiles.length !== 2) {
                showError('Debes seleccionar exactamente 2 archivos para comparar');
                return;
            }
            
            showLoading();
            const formData = new FormData();
            formData.append('file1', selectedFiles[0]);
            formData.append('file2', selectedFiles[1]);
            
            fetch('/compare', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    hideLoading();
                    if (data.error) {
                        showError(data.error);
                    } else {
                        displayComparison(data);
                    }
                })
                .catch(e => {
                    hideLoading();
                    showError('Error: ' + e.message);
                });
        }
        
        function displaySingleAnalysis(data) {
            const html = `
                <div class="results">
                    <h3>${data.file_name}</h3>
                    <div class="metrics-grid">
                        <div class="metric-card">
                            <h3>Filas</h3>
                            <div class="value">${data.rows}</div>
                        </div>
                        <div class="metric-card">
                            <h3>Columnas</h3>
                            <div class="value">${data.columns}</div>
                        </div>
                        <div class="metric-card">
                            <h3>Filas Duplicadas</h3>
                            <div class="value">${data.duplicates}</div>
                        </div>
                        <div class="metric-card">
                            <h3>Uso de RAM</h3>
                            <div class="value">${data.memory_usage} KB</div>
                        </div>
                    </div>
                    
                    <h4 style="margin-top: 25px; color: #333;">Columnas</h4>
                    <div class="schema-diff">
                        ${Object.entries(data.column_names).map(([idx, col]) => `
                            <div class="schema-item">
                                <h5>${col}</h5>
                                <p style="font-size: 0.85em; color: #666;">Tipo: ${data.dtypes[col]}</p>
                                <p style="font-size: 0.85em; color: #666;">Vacíos: ${data.missing_values[col]}</p>
                            </div>
                        `).join('')}
                    </div>
                    
                    ${Object.keys(data.numeric_stats || {}).length > 0 ? `
                        <h4 style="margin-top: 25px; color: #333;">Estadísticas Numéricas</h4>
                        <div class="schema-diff">
                            ${Object.entries(data.numeric_stats).map(([col, stats]) => `
                                <div class="schema-item">
                                    <h5>${col}</h5>
                                    <p style="font-size: 0.85em;">Min: ${stats.min}</p>
                                    <p style="font-size: 0.85em;">Media: ${stats.mean}</p>
                                    <p style="font-size: 0.85em;">Max: ${stats.max}</p>
                                </div>
                            `).join('')}
                        </div>
                    ` : ''}
                </div>
            `;
            document.getElementById('resultsContainer').innerHTML = html;
        }
        
        function displayComparison(data) {
            const schema = data.schema_differences;
            const dataDiff = data.data_differences;
            const metrics = data.metrics;
            
            const html = `
                <div class="results">
                    <h3>Comparación: ${data.file1} vs ${data.file2}</h3>
                    
                    <div class="metrics-grid">
                        <div class="metric-card">
                            <h3>Diferencia de Filas</h3>
                            <div class="value" style="color: ${metrics.row_difference !== 0 ? '#ff6b6b' : '#6bff6b'}">${metrics.row_difference > 0 ? '+' : ''}${metrics.row_difference}</div>
                        </div>
                        <div class="metric-card">
                            <h3>Diferencia de Columnas</h3>
                            <div class="value" style="color: ${metrics.col_difference !== 0 ? '#ff6b6b' : '#6bff6b'}">${metrics.col_difference > 0 ? '+' : ''}${metrics.col_difference}</div>
                        </div>
                        <div class="metric-card">
                            <h3>Columnas Diferentes</h3>
                            <div class="value" style="color: ${schema.total_different > 0 ? '#ff6b6b' : '#6bff6b'}">${schema.total_different}</div>
                        </div>
                        <div class="metric-card">
                            <h3>Columnas Comunes</h3>
                            <div class="value">${schema.common_count}</div>
                        </div>
                    </div>
                    
                    ${data.debug_samples && Object.keys(data.debug_samples).length > 0 ? `
                        <div class="differences">
                            <h4>🔍 Ejemplos de Diferencias Detectadas</h4>
                            <p style="margin-bottom: 15px; color: #666;">Primeros casos donde se detectaron diferencias:</p>
                            ${Object.entries(data.debug_samples).map(([col, samples]) => `
                                <div style="margin-bottom: 20px;">
                                    <h5 style="color: #667eea; margin-bottom: 10px;">Columna: ${col}</h5>
                                    ${samples.map((sample, idx) => `
                                        <div style="background: white; padding: 10px; margin: 8px 0; border-radius: 4px; border-left: 3px solid #ff6b6b;">
                                            <div style="font-size: 0.85em; color: #666; margin-bottom: 5px;">Fila ${sample.row + 2} (índice ${sample.row})</div>
                                            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                                                <div>
                                                    <strong style="color: #667eea;">Archivo 1:</strong>
                                                    <div style="font-family: monospace; background: #f0f2ff; padding: 5px; border-radius: 3px; margin-top: 3px;">
                                                        Valor: <code>${sample.file1_value}</code><br>
                                                        Tipo: ${sample.file1_type}<br>
                                                        Repr: <code style="font-size: 0.8em;">${sample.file1_repr}</code>
                                                    </div>
                                                </div>
                                                <div>
                                                    <strong style="color: #764ba2;">Archivo 2:</strong>
                                                    <div style="font-family: monospace; background: #f0f2ff; padding: 5px; border-radius: 3px; margin-top: 3px;">
                                                        Valor: <code>${sample.file2_value}</code><br>
                                                        Tipo: ${sample.file2_type}<br>
                                                        Repr: <code style="font-size: 0.8em;">${sample.file2_repr}</code>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    `).join('')}
                                </div>
                            `).join('')}
                        </div>
                    ` : ''}
                    
                    ${schema.only_in_file1.length > 0 ? `
                        <div class="differences">
                            <h4>❌ Solo en Archivo 1</h4>
                            <ul style="list-style-position: inside;">
                                ${schema.only_in_file1.map(c => '<li>' + c + '</li>').join('')}
                            </ul>
                        </div>
                    ` : ''}
                    
                    ${schema.only_in_file2.length > 0 ? `
                        <div class="differences">
                            <h4>❌ Solo en Archivo 2</h4>
                            <ul style="list-style-position: inside;">
                                ${schema.only_in_file2.map(c => '<li>' + c + '</li>').join('')}
                            </ul>
                        </div>
                    ` : ''}
                    
                    ${Object.keys(schema.type_changes).length > 0 ? `
                        <div class="differences">
                            <h4>⚠️ Cambios de Tipo</h4>
                            <div class="schema-diff">
                                ${Object.entries(schema.type_changes).map(([col, change]) => `
                                    <div class="schema-item">
                                        <h5>${col}</h5>
                                        <p>De: ${change.from}</p>
                                        <p>A: ${change.to}</p>
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                    ` : ''}
                    
                    ${Object.keys(dataDiff.column_differences).length > 0 ? `
                        <div class="differences">
                            <h4>📊 Diferencias en Datos</h4>
                            <div class="schema-diff">
                                ${Object.entries(dataDiff.column_differences).map(([col, count]) => `
                                    <div class="schema-item">
                                        <h5>${col}</h5>
                                        <p style="color: #ff6b6b; font-weight: bold;">${count} celdas diferentes</p>
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                    ` : '<div class="success">✅ Los datos son idénticos en columnas comunes</div>'}
                </div>
            `;
            document.getElementById('resultsContainer').innerHTML = html;
        }
        
        function showError(msg) {
            const html = `<div class="error"><strong>Error:</strong> ${msg}</div>`;
            document.getElementById('resultsContainer').innerHTML = html;
        }
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/analyze-single', methods=['POST'])
def analyze_single():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], 
                           secure_filename(file.filename))
    file.save(filepath)
    
    try:
        result = analyzer.analyze_single_file(filepath)
        return jsonify(result)
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

@app.route('/compare', methods=['POST'])
def compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': 'Se requieren 2 archivos'}), 400
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    path1 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file1.filename))
    path2 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file2.filename))
    
    file1.save(path1)
    file2.save(path2)
    
    try:
        result = analyzer.compare_files(path1, path2, debug_mode=True)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        for p in [path1, path2]:
            if os.path.exists(p):
                os.remove(p)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
