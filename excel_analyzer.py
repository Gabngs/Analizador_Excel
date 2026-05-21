import os
import json
import sqlite3
import tempfile
from pathlib import Path
from typing import Dict, List, Any
import pandas as pd
import numpy as np
from flask import Flask, render_template_string, request, jsonify, Response
from werkzeug.utils import secure_filename
import hashlib

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

MAX_DIFF_ROWS_STORED = 2000   # rows kept in SQLite / download report
MAX_DIFF_ROWS_UI = 200        # rows shown inline in the browser


class ExcelAnalyzer:
    def __init__(self):
        self.db_path = os.path.join(tempfile.gettempdir(), 'analyzer.db')
        self._init_db()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute(
            'CREATE TABLE IF NOT EXISTS comparisons (id TEXT PRIMARY KEY, data TEXT)'
        )
        conn.commit()
        conn.close()

    def get_sheet_info(self, file_path: str) -> Dict[str, Any]:
        try:
            xls = pd.ExcelFile(file_path)
            sheets = {}
            for name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=name)
                sheets[name] = {
                    'columns': df.columns.tolist(),
                    'dtypes': df.dtypes.astype(str).to_dict(),
                    'shape': list(df.shape),
                }
            return sheets
        except Exception as e:
            return {'error': str(e)}

    def analyze_single_file(self, file_path: str) -> Dict[str, Any]:
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
                'shape_info': f"{len(df)} filas × {len(df.columns)} columnas",
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

            try:
                xls = pd.ExcelFile(file_path)
                if len(xls.sheet_names) > 1:
                    metrics['sheets_warning'] = (
                        f"El archivo tiene {len(xls.sheet_names)} hojas: "
                        f"{', '.join(xls.sheet_names)}. Solo se analizó la primera."
                    )
                    metrics['sheet_names'] = xls.sheet_names
            except Exception:
                pass

            return metrics
        except Exception as e:
            return {'error': str(e)}

    # ------------------------------------------------------------------
    # Main comparison entry point
    # ------------------------------------------------------------------

    def compare_files(
        self, file1_path: str, file2_path: str, case_sensitive: bool = True
    ) -> Dict[str, Any]:
        try:
            df1 = pd.read_excel(file1_path)
            df2 = pd.read_excel(file2_path)

            sheets_info: Dict[str, List[str]] = {}
            for label, path in [('file1', file1_path), ('file2', file2_path)]:
                try:
                    with pd.ExcelFile(path) as xls:
                        sheets_info[label] = xls.sheet_names
                except Exception:
                    sheets_info[label] = []

            warnings: List[str] = []
            for label, sheets in sheets_info.items():
                if len(sheets) > 1:
                    fname = Path(file1_path if label == 'file1' else file2_path).name
                    warnings.append(
                        f"'{fname}' tiene {len(sheets)} hojas "
                        f"({', '.join(sheets)}). Solo se comparó la primera hoja."
                    )

            schema_diff = self._compare_schemas(df1, df2)
            data_diff = self._compare_data(df1, df2, case_sensitive)

            # Unique ID using content hash so the same files produce the same key
            comp_id = hashlib.md5(
                f"{Path(file1_path).name}|{Path(file2_path).name}|"
                f"{case_sensitive}|{pd.Timestamp.now().isoformat()}".encode()
            ).hexdigest()

            comparison = {
                'id': comp_id,
                'file1': Path(file1_path).name,
                'file2': Path(file2_path).name,
                'timestamp': pd.Timestamp.now().isoformat(),
                'case_sensitive': case_sensitive,
                'warnings': warnings,
                'sheets_info': sheets_info,
                'schema_differences': schema_diff,
                'data_differences': data_diff,
                'metrics': {
                    'file1_rows': len(df1),
                    'file2_rows': len(df2),
                    'file1_cols': len(df1.columns),
                    'file2_cols': len(df2.columns),
                    'row_difference': len(df2) - len(df1),
                    'col_difference': len(df2.columns) - len(df1.columns),
                },
            }

            conn = sqlite3.connect(self.db_path)
            conn.execute(
                'INSERT OR REPLACE INTO comparisons VALUES (?, ?)',
                (comp_id, json.dumps(comparison)),
            )
            conn.commit()
            conn.close()

            return comparison
        except Exception as e:
            return {'error': str(e)}

    # ------------------------------------------------------------------
    # Schema comparison
    # ------------------------------------------------------------------

    def _compare_schemas(self, df1: pd.DataFrame, df2: pd.DataFrame) -> Dict[str, Any]:
        cols1 = set(df1.columns)
        cols2 = set(df2.columns)
        return {
            'only_in_file1': sorted(cols1 - cols2),
            'only_in_file2': sorted(cols2 - cols1),
            'common_columns': sorted(cols1 & cols2),
            'type_changes': self._get_type_changes(df1, df2, cols1 & cols2),
            'common_count': len(cols1 & cols2),
            'total_different': len(cols1 ^ cols2),
        }

    def _get_type_changes(self, df1, df2, common_cols) -> Dict[str, Any]:
        changes = {}
        for col in common_cols:
            t1, t2 = str(df1[col].dtype), str(df2[col].dtype)
            if t1 != t2:
                changes[col] = {'from': t1, 'to': t2}
        return changes

    # ------------------------------------------------------------------
    # Value normalisation
    # ------------------------------------------------------------------

    def _normalize_value(self, val, case_sensitive: bool = True):
        try:
            if pd.isna(val):
                return None
        except (TypeError, ValueError):
            pass

        if isinstance(val, str):
            s = val.strip()
            return s if case_sensitive else s.lower()

        if isinstance(val, (int, float, np.integer, np.floating)):
            return float(val)

        if isinstance(val, (pd.Timestamp, np.datetime64)):
            return pd.Timestamp(val).isoformat()

        return val

    @staticmethod
    def _safe_isna(val) -> bool:
        try:
            return bool(pd.isna(val))
        except (TypeError, ValueError):
            return False

    @staticmethod
    def _to_display(val) -> str:
        if ExcelAnalyzer._safe_isna(val):
            return '(vacío)'
        return str(val)

    # ------------------------------------------------------------------
    # Data comparison — the core
    # ------------------------------------------------------------------

    def _compare_data(
        self, df1: pd.DataFrame, df2: pd.DataFrame, case_sensitive: bool = True
    ) -> Dict[str, Any]:
        common_cols = sorted(set(df1.columns) & set(df2.columns))

        if not common_cols:
            return {'error': 'No hay columnas comunes para comparar'}

        df1_c = df1[common_cols].reset_index(drop=True)
        df2_c = df2[common_cols].reset_index(drop=True)
        min_len = min(len(df1_c), len(df2_c))

        result: Dict[str, Any] = {
            'missing_in_file2': max(0, len(df1_c) - len(df2_c)),
            'extra_in_file2': max(0, len(df2_c) - len(df1_c)),
            'rows_compared': min_len,
            'rows_with_differences': 0,   # filled below — NOT min_len
            'column_differences': {},
            'diff_rows': [],              # up to MAX_DIFF_ROWS_STORED cells
            'diff_rows_total': 0,         # real total before cap
        }

        if min_len == 0:
            return result

        rows_with_diff: set = set()
        all_cells: List[Dict] = []

        for col in common_cols:
            s1 = df1_c[col].iloc[:min_len]
            s2 = df2_c[col].iloc[:min_len]

            # Float columns: use np.isclose so 0.1+0.2 == 0.3
            if s1.dtype in ('float64', 'float32') and s2.dtype in ('float64', 'float32'):
                diff_mask = pd.Series(
                    ~np.isclose(
                        s1.fillna(np.nan).values,
                        s2.fillna(np.nan).values,
                        rtol=1e-9,
                        atol=1e-9,
                        equal_nan=True,   # NaN == NaN → not a difference
                    ),
                    index=s1.index,
                )
            else:
                n1 = s1.apply(lambda v: self._normalize_value(v, case_sensitive))
                n2 = s2.apply(lambda v: self._normalize_value(v, case_sensitive))
                both_nan = n1.isna() & n2.isna()
                diff_mask = (n1 != n2) & ~both_nan

            diff_idx = diff_mask[diff_mask].index.tolist()
            if not diff_idx:
                continue

            result['column_differences'][col] = len(diff_idx)
            rows_with_diff.update(diff_idx)

            for idx in diff_idx:
                v1 = df1_c[col].iloc[idx]
                v2 = df2_c[col].iloc[idx]
                all_cells.append({
                    'row': int(idx),
                    'row_excel': int(idx) + 2,  # +1 header +1 to make 1-based
                    'col': str(col),
                    'val1': self._to_display(v1),
                    'val2': self._to_display(v2),
                    'null1': self._safe_isna(v1),
                    'null2': self._safe_isna(v2),
                })

        result['rows_with_differences'] = len(rows_with_diff)
        all_cells.sort(key=lambda x: (x['row'], x['col']))
        result['diff_rows_total'] = len(all_cells)
        result['diff_rows'] = all_cells[:MAX_DIFF_ROWS_STORED]

        return result


analyzer = ExcelAnalyzer()

# ──────────────────────────────────────────────────────────────────────────────
# HTML template
# ──────────────────────────────────────────────────────────────────────────────

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="es">
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
        .container { max-width: 1300px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 20px 60px rgba(0,0,0,.3); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p  { font-size: 1.1em; opacity: .9; }
        .content { padding: 40px; }
        .section { margin-bottom: 40px; }
        .section h2 { font-size: 1.5em; color: #333; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #667eea; }
        .upload-area { border: 2px dashed #667eea; border-radius: 8px; padding: 30px; text-align: center; cursor: pointer; transition: all .3s; background: #f8f9ff; }
        .upload-area:hover, .upload-area.dragover { border-color: #764ba2; background: #e8ebff; }
        input[type="file"] { display: none; }
        .file-input-label { cursor: pointer; color: #667eea; font-weight: 600; text-decoration: underline; }
        .button { display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 30px; border-radius: 6px; border: none; cursor: pointer; font-size: 1em; transition: transform .2s, box-shadow .2s; margin: 10px 5px 10px 0; text-decoration: none; }
        .button:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(102,126,234,.3); }
        .button:disabled { opacity: .6; cursor: not-allowed; transform: none; }
        .button-sm { padding: 8px 18px; font-size: .9em; }
        .button-green { background: linear-gradient(135deg, #43a047 0%, #1b5e20 100%); }
        .file-list { list-style: none; margin: 20px 0; }
        .file-item { background: #f8f9ff; padding: 12px; margin: 8px 0; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; }
        .file-item span { color: #333; font-weight: 500; }
        .file-remove { background: #ff6b6b; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; }
        .file-remove:hover { background: #ff5252; }
        .results { background: #f8f9ff; border-radius: 8px; padding: 25px; margin-top: 20px; }
        .metrics-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin: 20px 0; }
        .metric-card { background: white; border-left: 4px solid #667eea; padding: 15px; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
        .metric-card h3 { font-size: .85em; color: #666; margin-bottom: 8px; text-transform: uppercase; }
        .metric-card .value { font-size: 1.8em; color: #667eea; font-weight: bold; }
        .metric-card .value.red { color: #e53935; }
        .metric-card .value.green { color: #43a047; }
        .differences { background: white; border-radius: 6px; padding: 15px; margin: 15px 0; border-left: 4px solid #ff6b6b; }
        .differences h4 { color: #c62828; margin-bottom: 10px; }
        .warning-box { background: #fff8e1; border-left: 4px solid #f9a825; border-radius: 6px; padding: 12px 16px; margin: 12px 0; color: #6d4c00; font-size: .95em; }
        .schema-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 15px; margin: 15px 0; }
        .schema-item { background: white; padding: 12px; border-radius: 6px; border: 1px solid #e0e0e0; }
        .schema-item h5 { color: #555; margin-bottom: 8px; font-size: .85em; text-transform: uppercase; }
        .schema-item ul { list-style-position: inside; font-size: .9em; color: #444; }
        .schema-item li { padding: 3px 0; word-break: break-word; }
        .loading { display: none; text-align: center; padding: 30px; }
        .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #667eea; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 15px; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .error   { background: #ffe0e0; color: #c33; padding: 15px; border-radius: 6px; margin: 15px 0; border-left: 4px solid #ff6b6b; }
        .success { background: #e0ffe0; color: #2e7d32; padding: 15px; border-radius: 6px; margin: 15px 0; border-left: 4px solid #43a047; }
        /* Diff table */
        .diff-section { margin: 20px 0; }
        .diff-section h4 { color: #333; margin-bottom: 12px; font-size: 1.1em; }
        .diff-meta { font-size: .9em; color: #666; margin-bottom: 12px; display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
        .diff-table-wrap { overflow-x: auto; max-height: 480px; border: 1px solid #e0e0e0; border-radius: 6px; }
        .diff-table { width: 100%; border-collapse: collapse; font-size: .9em; }
        .diff-table thead th { background: #667eea; color: white; padding: 10px 14px; text-align: left; position: sticky; top: 0; z-index: 1; }
        .diff-table tbody tr:nth-child(odd) { background: #fafafa; }
        .diff-table tbody tr:hover { background: #f0f2ff; }
        .diff-table td { padding: 8px 14px; border-bottom: 1px solid #eee; max-width: 320px; word-break: break-word; vertical-align: top; }
        .diff-table td.val-null { color: #aaa; font-style: italic; }
        .diff-table td.val2 { background: #ffebee; }
        .diff-table .row-num { color: #888; font-size: .85em; font-family: monospace; }
        .col-badge { display: inline-block; background: #e8eafd; color: #3949ab; border-radius: 4px; padding: 2px 8px; font-size: .85em; font-weight: 600; }
        .options-row { display: flex; align-items: center; gap: 20px; flex-wrap: wrap; margin: 14px 0; }
        .checkbox-label { display: flex; align-items: center; gap: 6px; cursor: pointer; font-size: .95em; color: #444; }
        @media (max-width: 768px) { .metrics-grid { grid-template-columns: 1fr 1fr; } .header h1 { font-size: 1.8em; } .content { padding: 20px; } }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>&#128202; Excel Analyzer</h1>
        <p>Compara y analiza archivos Excel con precisión</p>
    </div>

    <div class="content">
        <div class="section">
            <h2>Cargar Archivos</h2>
            <div class="upload-area" id="uploadArea"
                 ondrop="handleDrop(event)" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)">
                <p>Arrastra archivos aquí o
                   <label class="file-input-label" for="fileInput">selecciona desde tu PC</label></p>
                <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" multiple onchange="handleFileSelect(event)">
            </div>

            <ul class="file-list" id="fileList"></ul>

            <div class="options-row">
                <label class="checkbox-label">
                    <input type="checkbox" id="caseSensitive" checked>
                    Comparación sensible a mayúsculas/minúsculas
                </label>
            </div>

            <div>
                <button class="button" onclick="analyzeSingle()" id="singleBtn" disabled>Analizar Archivo</button>
                <button class="button" onclick="doCompare()" id="compareBtn" disabled>Comparar 2 Archivos</button>
                <button class="button" onclick="clearFiles()" style="background:#999">Limpiar</button>
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

function handleDragOver(e) { e.preventDefault(); uploadArea.classList.add('dragover'); }
function handleDragLeave(e) { uploadArea.classList.remove('dragover'); }
function handleDrop(e) {
    e.preventDefault(); uploadArea.classList.remove('dragover');
    addFiles(Array.from(e.dataTransfer.files));
}
function handleFileSelect(e) { addFiles(Array.from(e.target.files)); }

function addFiles(files) {
    selectedFiles = [...selectedFiles, ...files].slice(0, 2);
    updateFileList();
}
function removeFile(idx) { selectedFiles.splice(idx, 1); updateFileList(); }
function clearFiles() { selectedFiles = []; updateFileList(); document.getElementById('resultsContainer').innerHTML = ''; }

function updateFileList() {
    const list = document.getElementById('fileList');
    list.innerHTML = '';
    selectedFiles.forEach((f, i) => {
        const li = document.createElement('li');
        li.className = 'file-item';
        li.innerHTML = `<span>${f.name} (${(f.size/1024).toFixed(1)} KB)</span>
                        <button class="file-remove" onclick="removeFile(${i})">Eliminar</button>`;
        list.appendChild(li);
    });
    document.getElementById('singleBtn').disabled  = selectedFiles.length === 0;
    document.getElementById('compareBtn').disabled = selectedFiles.length !== 2;
}

function showLoading()  { document.getElementById('loading').style.display = 'block'; }
function hideLoading()  { document.getElementById('loading').style.display = 'none'; }
function showError(msg) { document.getElementById('resultsContainer').innerHTML = `<div class="error"><strong>Error:</strong> ${esc(msg)}</div>`; }

function esc(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function analyzeSingle() {
    if (!selectedFiles.length) return showError('Selecciona al menos un archivo');
    showLoading();
    const fd = new FormData(); fd.append('file', selectedFiles[0]);
    fetch('/analyze-single', { method:'POST', body:fd })
        .then(r => r.json())
        .then(d => { hideLoading(); d.error ? showError(d.error) : displaySingleAnalysis(d); })
        .catch(e => { hideLoading(); showError(e.message); });
}

function doCompare() {
    if (selectedFiles.length !== 2) return showError('Selecciona exactamente 2 archivos');
    showLoading();
    const fd = new FormData();
    fd.append('file1', selectedFiles[0]);
    fd.append('file2', selectedFiles[1]);
    fd.append('case_sensitive', document.getElementById('caseSensitive').checked ? '1' : '0');
    fetch('/compare', { method:'POST', body:fd })
        .then(r => r.json())
        .then(d => { hideLoading(); d.error ? showError(d.error) : displayComparison(d); })
        .catch(e => { hideLoading(); showError(e.message); });
}

function displaySingleAnalysis(d) {
    let cols = d.column_names.map(col => `
        <div class="schema-item">
            <h5>${esc(col)}</h5>
            <p style="font-size:.85em;color:#666">Tipo: ${esc(d.dtypes[col])}</p>
            <p style="font-size:.85em;color:#666">Vacíos: ${d.missing_values[col]}</p>
        </div>`).join('');

    let numStats = '';
    if (d.numeric_stats && Object.keys(d.numeric_stats).length) {
        numStats = '<h4 style="margin-top:25px;color:#333">Estadísticas Numéricas</h4><div class="schema-grid">'
            + Object.entries(d.numeric_stats).map(([col,s]) => `
                <div class="schema-item">
                    <h5>${esc(col)}</h5>
                    <p style="font-size:.85em">Min: ${s.min}</p>
                    <p style="font-size:.85em">Media: ${s.mean}</p>
                    <p style="font-size:.85em">Max: ${s.max}</p>
                </div>`).join('') + '</div>';
    }

    document.getElementById('resultsContainer').innerHTML = `
        <div class="results">
            <h3>${esc(d.file_name)}</h3>
            ${d.sheets_warning ? `<div class="warning-box">&#9888; ${esc(d.sheets_warning)}</div>` : ''}
            <div class="metrics-grid">
                <div class="metric-card"><h3>Filas</h3><div class="value">${d.rows}</div></div>
                <div class="metric-card"><h3>Columnas</h3><div class="value">${d.columns}</div></div>
                <div class="metric-card"><h3>Duplicados</h3><div class="value ${d.duplicates>0?'red':'green'}">${d.duplicates}</div></div>
                <div class="metric-card"><h3>RAM</h3><div class="value">${d.memory_usage} KB</div></div>
            </div>
            <h4 style="margin-top:25px;color:#333">Columnas</h4>
            <div class="schema-grid">${cols}</div>
            ${numStats}
        </div>`;
}

function displayComparison(data) {
    const schema   = data.schema_differences;
    const dataDiff = data.data_differences;
    const metrics  = data.metrics;

    // ── Summary cards ──────────────────────────────────────────────
    const rowDiffColor   = metrics.row_difference  !== 0 ? 'red' : 'green';
    const colDiffColor   = metrics.col_difference  !== 0 ? 'red' : 'green';
    const schemaDiffColor= schema.total_different  >  0  ? 'red' : 'green';
    const rowsDiffColor  = (dataDiff.rows_with_differences||0) > 0 ? 'red' : 'green';

    const caseNote = data.case_sensitive
        ? '<small style="color:#888">Modo: sensible a mayúsculas</small>'
        : '<small style="color:#e65100">Modo: ignora mayúsculas</small>';

    let html = `<div class="results">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px">
            <h3 style="color:#333">Comparación: <em>${esc(data.file1)}</em> vs <em>${esc(data.file2)}</em></h3>
            ${caseNote}
        </div>`;

    // Warnings
    (data.warnings||[]).forEach(w => {
        html += `<div class="warning-box">&#9888; ${esc(w)}</div>`;
    });

    html += `<div class="metrics-grid">
        <div class="metric-card"><h3>Filas Archivo 1</h3><div class="value">${metrics.file1_rows}</div></div>
        <div class="metric-card"><h3>Filas Archivo 2</h3><div class="value">${metrics.file2_rows}</div></div>
        <div class="metric-card"><h3>Dif. de Filas</h3><div class="value ${rowDiffColor}">${metrics.row_difference>0?'+':''}${metrics.row_difference}</div></div>
        <div class="metric-card"><h3>Dif. de Columnas</h3><div class="value ${colDiffColor}">${metrics.col_difference>0?'+':''}${metrics.col_difference}</div></div>
        <div class="metric-card"><h3>Cols. del Schema Distintas</h3><div class="value ${schemaDiffColor}">${schema.total_different}</div></div>
        <div class="metric-card"><h3>Filas con Diferencias</h3><div class="value ${rowsDiffColor}">${dataDiff.rows_with_differences||0}</div></div>
    </div>`;

    // ── Schema differences ──────────────────────────────────────────
    if (schema.only_in_file1.length || schema.only_in_file2.length || Object.keys(schema.type_changes).length) {
        html += `<div class="differences"><h4>&#10060; Diferencias en Esquema</h4><div class="schema-grid">`;

        if (schema.only_in_file1.length)
            html += `<div class="schema-item"><h5>Solo en Archivo 1</h5><ul>${schema.only_in_file1.map(c=>`<li>${esc(c)}</li>`).join('')}</ul></div>`;

        if (schema.only_in_file2.length)
            html += `<div class="schema-item"><h5>Solo en Archivo 2</h5><ul>${schema.only_in_file2.map(c=>`<li>${esc(c)}</li>`).join('')}</ul></div>`;

        if (Object.keys(schema.type_changes).length)
            html += `<div class="schema-item"><h5>Cambio de Tipo</h5><ul>`
                + Object.entries(schema.type_changes).map(([c,t])=>`<li>${esc(c)}: ${esc(t.from)} → ${esc(t.to)}</li>`).join('')
                + `</ul></div>`;

        html += `</div></div>`;
    }

    // ── Column-level counts ─────────────────────────────────────────
    const colDiffs = dataDiff.column_differences || {};
    if (Object.keys(colDiffs).length) {
        html += `<div class="differences"><h4>&#128202; Celdas Diferentes por Columna</h4><div class="schema-grid">`;
        Object.entries(colDiffs).sort((a,b)=>b[1]-a[1]).forEach(([col,cnt]) => {
            const pct = dataDiff.rows_compared
                ? ((cnt/dataDiff.rows_compared)*100).toFixed(1)
                : '?';
            html += `<div class="schema-item">
                <h5>${esc(col)}</h5>
                <p style="color:#e53935;font-weight:bold">${cnt} celdas (${pct}%)</p>
                <p style="font-size:.8em;color:#888">de ${dataDiff.rows_compared} filas comparadas</p>
            </div>`;
        });
        html += `</div></div>`;
    }

    // ── Full diff table ─────────────────────────────────────────────
    const diffRows = dataDiff.diff_rows || [];
    if (diffRows.length) {
        const totalCells = dataDiff.diff_rows_total || diffRows.length;
        const shown      = Math.min(diffRows.length, ${MAX_DIFF_ROWS_UI});
        const truncNote  = (totalCells > shown)
            ? `<span style="color:#e65100">Mostrando ${shown} de ${totalCells} celdas diferentes. Descarga el informe para verlas todas.</span>`
            : `<span style="color:#666">${totalCells} celdas diferentes</span>`;

        html += `<div class="diff-section">
            <h4>&#128269; Tabla de Diferencias</h4>
            <div class="diff-meta">
                ${truncNote}
                <a class="button button-sm button-green" href="/download-report?id=${esc(data.id)}" target="_blank">
                    &#8681; Descargar Informe HTML
                </a>
            </div>
            <div class="diff-table-wrap">
            <table class="diff-table">
                <thead>
                    <tr>
                        <th>Fila (Excel)</th>
                        <th>Columna</th>
                        <th>${esc(data.file1)}</th>
                        <th>${esc(data.file2)}</th>
                    </tr>
                </thead>
                <tbody>`;

        diffRows.slice(0, shown).forEach(r => {
            const cls1 = r.null1 ? ' val-null' : '';
            const cls2 = r.null2 ? ' val-null' : '';
            html += `<tr>
                <td class="row-num">${r.row_excel}</td>
                <td><span class="col-badge">${esc(r.col)}</span></td>
                <td class="${cls1}">${esc(r.val1)}</td>
                <td class="val2${cls2}">${esc(r.val2)}</td>
            </tr>`;
        });

        html += `</tbody></table></div></div>`;

    } else if (!Object.keys(colDiffs).length && !schema.total_different
               && !metrics.row_difference && !metrics.col_difference) {
        html += `<div class="success">&#10003; Los archivos son idénticos (mismas filas, columnas y datos)</div>`;
    } else if (!Object.keys(colDiffs).length) {
        html += `<div class="success">&#10003; Los datos en columnas comunes son idénticos</div>`;
    }

    // ── Rows count mismatch notice ──────────────────────────────────
    if (dataDiff.missing_in_file2 > 0)
        html += `<div class="warning-box">&#9888; ${dataDiff.missing_in_file2} filas presentes en Archivo 1 no existen en Archivo 2 (o viceversa — archivo más corto)</div>`;
    if (dataDiff.extra_in_file2 > 0)
        html += `<div class="warning-box">&#9888; ${dataDiff.extra_in_file2} filas extra en Archivo 2 (no comparadas)</div>`;

    html += `</div>`;
    document.getElementById('resultsContainer').innerHTML = html;
}
</script>
</body>
</html>
'''.replace('${MAX_DIFF_ROWS_UI}', str(MAX_DIFF_ROWS_UI))


# ──────────────────────────────────────────────────────────────────────────────
# Download report generator
# ──────────────────────────────────────────────────────────────────────────────

def _report_schema_rows(schema: dict) -> str:
    rows = ''
    for col in schema.get('only_in_file1', []):
        rows += f'<tr><td>{col}</td><td class="y">Solo en Archivo 1</td><td>—</td></tr>'
    for col in schema.get('only_in_file2', []):
        rows += f'<tr><td>{col}</td><td class="y">Solo en Archivo 2</td><td>—</td></tr>'
    for col, chg in schema.get('type_changes', {}).items():
        rows += f'<tr><td>{col}</td><td class="y">Cambio de tipo</td><td>{chg["from"]} → {chg["to"]}</td></tr>'
    return rows


def _report_col_rows(col_diffs: dict, rows_compared: int) -> str:
    rows = ''
    for col, cnt in sorted(col_diffs.items(), key=lambda x: -x[1]):
        pct = f'{cnt / rows_compared * 100:.1f}%' if rows_compared else '?'
        rows += f'<tr><td>{col}</td><td class="diff">{cnt}</td><td>{pct}</td></tr>'
    return rows


def _report_data_rows(diff_rows: list) -> str:
    rows = ''
    for r in diff_rows:
        n1 = ' class="null"' if r.get('null1') else ''
        n2 = ' class="null diff"' if r.get('null2') else ' class="diff"'
        rows += (
            f'<tr>'
            f'<td class="mono">{r["row_excel"]}</td>'
            f'<td><b>{r["col"]}</b></td>'
            f'<td{n1}>{r["val1"]}</td>'
            f'<td{n2}>{r["val2"]}</td>'
            f'</tr>'
        )
    return rows


def _report_summary_cards(metrics: dict, schema: dict, dd: dict, total_cells: int) -> str:
    def color(bad_condition: bool) -> str:
        return 'red' if bad_condition else 'green'

    row_diff = metrics.get('row_difference', 0)
    return f'''<div class="summary">
  <div class="card"><div class="lbl">Filas Archivo 1</div><div class="val">{metrics.get("file1_rows", 0)}</div></div>
  <div class="card"><div class="lbl">Filas Archivo 2</div><div class="val">{metrics.get("file2_rows", 0)}</div></div>
  <div class="card"><div class="lbl">Dif. Filas</div><div class="val {color(row_diff != 0)}">{row_diff:+}</div></div>
  <div class="card"><div class="lbl">Cols. Comunes</div><div class="val">{schema.get("common_count", 0)}</div></div>
  <div class="card"><div class="lbl">Cols. Diferentes</div><div class="val {color(schema.get("total_different", 0) > 0)}">{schema.get("total_different", 0)}</div></div>
  <div class="card"><div class="lbl">Filas con Diffs</div><div class="val {color(dd.get("rows_with_differences", 0) > 0)}">{dd.get("rows_with_differences", 0)}</div></div>
  <div class="card"><div class="lbl">Celdas Diferentes</div><div class="val {color(total_cells > 0)}">{total_cells}</div></div>
</div>'''


def _build_download_report(data: dict) -> str:
    schema    = data.get('schema_differences', {})
    dd        = data.get('data_differences', {})
    metrics   = data.get('metrics', {})
    diff_rows = dd.get('diff_rows', [])
    col_diffs = dd.get('column_differences', {})

    schema_rows = _report_schema_rows(schema)
    col_rows    = _report_col_rows(col_diffs, dd.get('rows_compared', 0))
    data_rows   = _report_data_rows(diff_rows)

    warnings_html = ''.join(
        f'<p class="warn">&#9888; {w}</p>' for w in data.get('warnings', [])
    )

    total_cells = dd.get('diff_rows_total', len(diff_rows))
    truncation_note = (
        f'<p class="warn">Nota: se muestran {len(diff_rows)} de {total_cells} celdas diferentes (límite del informe).</p>'
        if total_cells > len(diff_rows) else ''
    )

    summary_cards = _report_summary_cards(metrics, schema, dd, total_cells)

    return f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Informe de Comparación — {data.get("file1","?")} vs {data.get("file2","?")}</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 30px; background: #f9f9f9; color: #222; }}
  h1 {{ color: #3949ab; }}
  h2 {{ color: #555; border-bottom: 2px solid #667eea; padding-bottom: 6px; margin: 30px 0 14px; }}
  .meta {{ background: #e8eaf6; border-radius: 6px; padding: 14px 20px; margin-bottom: 20px; font-size: .95em; }}
  .meta p {{ margin: 4px 0; }}
  .summary {{ display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 20px; }}
  .card {{ background: white; border-left: 4px solid #667eea; border-radius: 6px; padding: 14px 20px; min-width: 140px; box-shadow: 0 1px 4px rgba(0,0,0,.1); }}
  .card .lbl {{ font-size: .8em; color: #888; text-transform: uppercase; }}
  .card .val {{ font-size: 1.7em; font-weight: bold; color: #3949ab; }}
  .card .val.red {{ color: #c62828; }}
  .card .val.green {{ color: #2e7d32; }}
  table {{ width: 100%; border-collapse: collapse; background: white; font-size: .92em; margin-bottom: 16px; }}
  th {{ background: #3949ab; color: white; padding: 9px 12px; text-align: left; }}
  td {{ padding: 8px 12px; border-bottom: 1px solid #eee; }}
  tr:nth-child(even) {{ background: #f5f5f5; }}
  td.diff {{ background: #ffebee; color: #c62828; font-weight: 600; }}
  td.null {{ color: #aaa; font-style: italic; }}
  td.y    {{ color: #e65100; font-weight: 600; }}
  td.mono {{ font-family: monospace; color: #888; }}
  .warn {{ background: #fff8e1; border-left: 4px solid #f9a825; padding: 10px 14px; border-radius: 4px; margin: 8px 0; color: #6d4c00; }}
  .ok   {{ background: #e8f5e9; border-left: 4px solid #43a047; padding: 10px 14px; border-radius: 4px; margin: 8px 0; color: #1b5e20; }}
  @media print {{ body {{ margin: 10px; }} }}
</style>
</head>
<body>
<h1>&#128202; Informe de Comparación Excel</h1>

<div class="meta">
  <p><strong>Archivo 1:</strong> {data.get("file1","?")}</p>
  <p><strong>Archivo 2:</strong> {data.get("file2","?")}</p>
  <p><strong>Fecha:</strong> {data.get("timestamp","?")}</p>
  <p><strong>Sensible a mayúsculas:</strong> {"Sí" if data.get("case_sensitive") else "No"}</p>
</div>

{warnings_html}

<h2>Resumen</h2>
{summary_cards}

{"" if not schema_rows else f"""
<h2>Diferencias de Esquema</h2>
<table><thead><tr><th>Columna</th><th>Diferencia</th><th>Detalle</th></tr></thead>
<tbody>{schema_rows}</tbody></table>"""}

{"" if not col_rows else f"""
<h2>Celdas Diferentes por Columna</h2>
<table><thead><tr><th>Columna</th><th>Celdas Diferentes</th><th>% sobre filas comparadas</th></tr></thead>
<tbody>{col_rows}</tbody></table>"""}

{"" if not data_rows else f"""
<h2>Detalle Celda por Celda</h2>
{truncation_note}
<table>
  <thead>
    <tr>
      <th>Fila (Excel)</th>
      <th>Columna</th>
      <th>{data.get("file1","Archivo 1")}</th>
      <th>{data.get("file2","Archivo 2")} (diferente)</th>
    </tr>
  </thead>
  <tbody>{data_rows}</tbody>
</table>"""}

{"<div class='ok'>&#10003; Los datos en columnas comunes son idénticos.</div>" if not col_rows and not schema_rows else ""}

</body>
</html>'''


# ──────────────────────────────────────────────────────────────────────────────
# Flask routes
# ──────────────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/analyze-single', methods=['POST'])
def analyze_single():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    f = request.files['file']
    if not f.filename:
        return jsonify({'error': 'No file selected'}), 400

    path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename))
    f.save(path)
    try:
        return jsonify(analyzer.analyze_single_file(path))
    finally:
        if os.path.exists(path):
            os.remove(path)


@app.route('/compare', methods=['POST'])
def compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': 'Se requieren 2 archivos'}), 400

    f1 = request.files['file1']
    f2 = request.files['file2']
    case_sensitive = request.form.get('case_sensitive', '1') != '0'

    p1 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f1.filename))
    p2 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f2.filename))

    f1.save(p1)
    f2.save(p2)
    try:
        result = analyzer.compare_files(p1, p2, case_sensitive=case_sensitive)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # Dar tiempo para que se liberen los recursos del archivo
        import time
        import gc
        gc.collect()  # Forzar recolección de basura para cerrar referencias
        time.sleep(0.1)  # Pequeño delay para permitir que Windows libere el archivo
        
        for p in (p1, p2):
            if os.path.exists(p):
                # Intentar eliminar con reintentos en caso de bloqueo temporal
                max_attempts = 3
                for attempt in range(max_attempts):
                    try:
                        os.remove(p)
                        break
                    except PermissionError:
                        if attempt < max_attempts - 1:
                            time.sleep(0.5)  # Esperar antes de reintentar
                        else:
                            # Si falla después de todos los intentos, registrar pero no fallar
                            print(f"Advertencia: No se pudo eliminar {p}")


@app.route('/download-report', methods=['GET'])
def download_report():
    comp_id = request.args.get('id', '')
    if not comp_id:
        return 'ID de comparación requerido', 400

    conn = sqlite3.connect(analyzer.db_path)
    row = conn.execute(
        'SELECT data FROM comparisons WHERE id = ?', (comp_id,)
    ).fetchone()
    conn.close()

    if not row:
        return 'Comparación no encontrada. Realiza la comparación de nuevo.', 404

    data = json.loads(row[0])
    html = _build_download_report(data)

    fname = f"informe_{data.get('file1','?')}_vs_{data.get('file2','?')}.html"
    fname = fname.replace(' ', '_')

    return Response(
        html,
        mimetype='text/html',
        headers={'Content-Disposition': f'attachment; filename="{fname}"'},
    )


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
