import os
import re
import sys
import json
import sqlite3
import tempfile
import threading
import uuid
import webbrowser
from pathlib import Path
from typing import Dict, List, Any
import pandas as pd
import numpy as np
from flask import Flask, render_template_string, request, jsonify, Response
from werkzeug.utils import secure_filename
import hashlib

# Fix UnicodeEncodeError on Windows consoles that use cp1252 encoding
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# pandas dtype aliases treated as equivalent for comparison purposes
_DTYPE_EQUIV: Dict[str, str] = {'str': 'object', 'string': 'object'}

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
# Convertir la ruta corta de Windows a ruta larga para evitar problemas
app.config['UPLOAD_FOLDER'] = os.path.realpath(tempfile.gettempdir())
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

MAX_DIFF_ROWS_STORED = 2000   # rows kept in SQLite / download report
MAX_DIFF_ROWS_UI = 200        # rows shown inline in the browser


class ExcelAnalyzer:
    def __init__(self):
        # Usar ruta completa para evitar problemas con rutas cortas de Windows
        temp_dir = os.path.realpath(tempfile.gettempdir())
        self.db_path = os.path.join(temp_dir, 'analyzer.db')
        self._init_db()
    
    def _preprocess_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Preprocesa el DataFrame para corregir problemas comunes."""
        # Detect when pandas read a title row as column names (>50% "Unnamed" cols)
        # and the real headers are in the first data row.
        if len(df) > 0 and sum(1 for c in df.columns if 'Unnamed' in str(c)) > len(df.columns) * 0.5:
            try:
                first_row = df.iloc[0]
                if all(isinstance(v, str) or pd.isna(v) for v in first_row[:10]):
                    df = df.copy()
                    df.columns = first_row.fillna('').astype(str)
                    df = df.iloc[1:].reset_index(drop=True)
            except Exception as e:
                print(f"Warning: could not promote first row to header: {e}")

        # Strip timestamps embedded in column names and trim whitespace
        new_cols = []
        seen: Dict[str, int] = {}
        for col in df.columns:
            col_str = str(col).strip()
            col_str = re.sub(r'\d{4}-\d{2}-\d{2}[\s_]\d{2}[:-]\d{2}[:-]\d{2}', '', col_str)
            col_str = re.sub(r'\d{4}-\d{2}-\d{2}', '', col_str)
            col_str = col_str.strip(' -_,')
            if not col_str:
                col_str = f'COL_{len(new_cols)}'
            # Deduplicate column names
            if col_str in seen:
                seen[col_str] += 1
                col_str = f'{col_str}_{seen[col_str]}'
            else:
                seen[col_str] = 0
            new_cols.append(col_str)
        df.columns = new_cols

        # Normalize leading/trailing whitespace in text columns
        for col in df.select_dtypes(include=['object']).columns:
            try:
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            except Exception:
                pass

        return df

    # ------------------------------------------------------------------
    # Smart Excel reader — auto-detects the real header row
    # ------------------------------------------------------------------

    def _find_header_row(self, df_raw: pd.DataFrame) -> int:
        """Return the 0-based row index that best serves as column headers.

        Scans the first 20 rows looking for a row where all non-empty cells
        are strings (text labels), the row is well-filled, and the next row
        actually contains data.  Falls back to 0 if nothing scores well enough.
        """
        total_cols = len(df_raw.columns)
        max_check = min(20, max(0, len(df_raw) - 1))

        best_row = 0
        best_score = -1.0

        for i in range(max_check):
            row_vals = df_raw.iloc[i].tolist()
            non_null = [v for v in row_vals if pd.notna(v) and str(v).strip() != '']

            if len(non_null) < max(2, total_cols * 0.15):
                continue  # too sparse to be a header

            if not all(isinstance(v, str) for v in non_null):
                continue  # real header rows contain only text labels

            fill_ratio   = len(non_null) / total_cols
            unique_ratio = len(set(non_null)) / len(non_null)

            has_data_after = False
            for j in range(i + 1, min(i + 3, len(df_raw))):
                next_row = [v for v in df_raw.iloc[j].tolist()
                            if pd.notna(v) and str(v).strip() != '']
                if len(next_row) >= len(non_null) * 0.4:
                    has_data_after = True
                    break

            score = fill_ratio * 0.5 + unique_ratio * 0.3 + (0.2 if has_data_after else 0.0)

            if score > best_score:
                best_score = score
                best_row = i

        return best_row

    def _smart_read_excel(self, file_path: str) -> pd.DataFrame:
        """Read an Excel file with automatic header-row detection.

        Reads the first 30 rows raw to find the real header, then re-reads
        the full file with the correct ``header`` parameter for proper dtype
        inference.
        """
        df_raw = pd.read_excel(file_path, header=None, nrows=30)
        header_row = self._find_header_row(df_raw)
        return pd.read_excel(file_path, header=header_row)

    @staticmethod
    def _clean_display_name(file_path: str) -> str:
        """Remove the UUID temp prefix added by Flask routes from a filename."""
        name = Path(file_path).name
        # Pattern: 32 hex chars + optional _1_ or _2_ separator
        return re.sub(r'^[0-9a-f]{32}_\d?_?', '', name)

    def _init_db(self):
        """Inicializa la base de datos SQLite, verificando que la tabla existe."""
        try:
            conn = sqlite3.connect(self.db_path)
            conn.execute(
                'CREATE TABLE IF NOT EXISTS comparisons (id TEXT PRIMARY KEY, data TEXT)'
            )
            conn.commit()
            conn.close()
            print(f"✅ Base de datos inicializada: {self.db_path}")
        except Exception as e:
            print(f"❌ Error al inicializar base de datos: {e}")
            raise
    
    def _ensure_db(self):
        """Verifica que la base de datos y la tabla existan antes de usarlas."""
        try:
            conn = sqlite3.connect(self.db_path)
            # Verificar si la tabla existe
            cursor = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='comparisons'"
            )
            if cursor.fetchone() is None:
                # La tabla no existe, crearla
                conn.execute(
                    'CREATE TABLE comparisons (id TEXT PRIMARY KEY, data TEXT)'
                )
                conn.commit()
            conn.close()
        except Exception as e:
            print(f"⚠️ Error al verificar base de datos: {e}")
            # Intentar recrear la base de datos
            self._init_db()

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

    def analyze_single_file(self, file_path: str, display_name: str = None) -> Dict[str, Any]:
        try:
            df = self._smart_read_excel(file_path)
            df = self._preprocess_dataframe(df)

            metrics = {
                'file_name': display_name or self._clean_display_name(file_path),
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
        self, file1_path: str, file2_path: str, case_sensitive: bool = True,
        sort_columns: List[str] = None, auto_sort: bool = True, ignore_type_diff: bool = True,
        file1_name: str = None, file2_name: str = None,
    ) -> Dict[str, Any]:
        try:
            df1 = self._smart_read_excel(file1_path)
            df2 = self._smart_read_excel(file2_path)

            df1 = self._preprocess_dataframe(df1)
            df2 = self._preprocess_dataframe(df2)

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
            # Si ignore_type_diff está activo, no mostrar cambios de tipo
            if ignore_type_diff:
                schema_diff['type_changes'] = {}
            data_diff = self._compare_data(df1, df2, case_sensitive, sort_columns, auto_sort, ignore_type_diff)

            # Unique ID using content hash so the same files produce the same key
            comp_id = hashlib.md5(
                f"{Path(file1_path).name}|{Path(file2_path).name}|"
                f"{case_sensitive}|{pd.Timestamp.now().isoformat()}".encode()
            ).hexdigest()

            comparison = {
                'id': comp_id,
                'file1': file1_name or self._clean_display_name(file1_path),
                'file2': file2_name or self._clean_display_name(file2_path),
                'timestamp': pd.Timestamp.now().isoformat(),
                'case_sensitive': case_sensitive,
                'ignore_type_diff': ignore_type_diff,
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

            # Asegurar que la base de datos y tabla existen
            self._ensure_db()
            
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
            t1_raw, t2_raw = str(df1[col].dtype), str(df2[col].dtype)
            # Normalize equivalent dtype names before comparing
            t1 = _DTYPE_EQUIV.get(t1_raw, t1_raw)
            t2 = _DTYPE_EQUIV.get(t2_raw, t2_raw)
            if t1 != t2:
                changes[col] = {'from': t1_raw, 'to': t2_raw}
        return changes

    # ------------------------------------------------------------------
    # Detect candidate columns for sorting
    # ------------------------------------------------------------------

    def _detect_sort_columns(self, df1: pd.DataFrame, df2: pd.DataFrame) -> List[str]:
        """Detecta automáticamente columnas candidatas para ordenar antes de comparar.
        Prioriza: ID, código, fecha, nombre, índice único.
        """
        common_cols = list(set(df1.columns) & set(df2.columns))
        if not common_cols:
            return []

        candidates = []
        
        # Palabras clave que indican columnas de identificación/ordenamiento
        key_patterns = ['id', 'codigo', 'code', 'clave', 'key', 'index', 'indice',
                       'num', 'numero', 'number', 'folio', 'orden', 'order']
        
        # Buscar columnas con nombres que sugieran identificadores
        for col in common_cols:
            col_lower = str(col).lower()
            if any(pattern in col_lower for pattern in key_patterns):
                # Verificar que la columna sea útil para ordenar (sin muchos duplicados)
                try:
                    df1_unique_ratio = df1[col].nunique() / len(df1) if len(df1) > 0 else 0
                    df2_unique_ratio = df2[col].nunique() / len(df2) if len(df2) > 0 else 0
                    # Si tiene al menos 50% de valores únicos, es buena candidata
                    if df1_unique_ratio >= 0.5 and df2_unique_ratio >= 0.5:
                        candidates.append(col)
                except (TypeError, ValueError, KeyError):
                    pass
        
        # Si no encontramos candidatos por nombre, buscar columnas con alta cardinalidad
        if not candidates:
            for col in common_cols:
                try:
                    # Verificar que sea ordenable (numérico, string, o datetime)
                    if df1[col].dtype in ['int64', 'float64', 'object', 'datetime64[ns]']:
                        df1_unique_ratio = df1[col].nunique() / len(df1) if len(df1) > 0 else 0
                        df2_unique_ratio = df2[col].nunique() / len(df2) if len(df2) > 0 else 0
                        if df1_unique_ratio >= 0.7 and df2_unique_ratio >= 0.7:
                            candidates.append((col, df1_unique_ratio))
                except (TypeError, ValueError, KeyError):
                    pass
            
            # Ordenar por cardinalidad y tomar las mejores
            if candidates and isinstance(candidates[0], tuple):
                candidates.sort(key=lambda x: x[1], reverse=True)
                candidates = [col for col, _ in candidates[:3]]
        
        return candidates[:3]  # Máximo 3 columnas para ordenar

    # ------------------------------------------------------------------
    # Value normalisation
    # ------------------------------------------------------------------

    def _normalize_value(self, val, case_sensitive: bool = True, ignore_type_diff: bool = False):
        try:
            if pd.isna(val):
                return None
        except (TypeError, ValueError):
            pass

        if isinstance(val, str):
            s = val.strip()
            # Strip trailing decimal dot — Excel export artifact: '12345.' → '12345'
            if s.endswith('.') and s[:-1].lstrip('-').isdigit():
                s = s[:-1]
            if ignore_type_diff:
                try:
                    return float(s)
                except (ValueError, TypeError):
                    return s if case_sensitive else s.lower()
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
        self, df1: pd.DataFrame, df2: pd.DataFrame, case_sensitive: bool = True,
        sort_columns: List[str] = None, auto_sort: bool = True, ignore_type_diff: bool = False
    ) -> Dict[str, Any]:
        common_cols = sorted(set(df1.columns) & set(df2.columns))

        if not common_cols:
            return {'error': 'No hay columnas comunes para comparar'}

        df1_c = df1[common_cols].copy()
        df2_c = df2[common_cols].copy()
        
        # Determinar columnas para ordenamiento
        sort_by = []
        sort_info = {'sorted': False, 'columns': []}
        
        # Solo ordenar si auto_sort está activado
        if auto_sort:
            if sort_columns:
                # Usar columnas especificadas por el usuario
                sort_by = [col for col in sort_columns if col in common_cols]
            else:
                # Auto-detectar columnas para ordenar
                sort_by = self._detect_sort_columns(df1_c, df2_c)
        
        # Aplicar ordenamiento si hay columnas identificadas
        if sort_by:
            try:
                # Ordenar ambos DataFrames por las mismas columnas
                df1_c = df1_c.sort_values(by=sort_by, na_position='last').reset_index(drop=True)
                df2_c = df2_c.sort_values(by=sort_by, na_position='last').reset_index(drop=True)
                sort_info = {'sorted': True, 'columns': sort_by}
            except Exception as e:
                # Si falla el ordenamiento, continuar sin ordenar
                sort_info = {'sorted': False, 'columns': sort_by, 'error': str(e)}
                df1_c = df1_c.reset_index(drop=True)
                df2_c = df2_c.reset_index(drop=True)
        else:
            df1_c = df1_c.reset_index(drop=True)
            df2_c = df2_c.reset_index(drop=True)
        
        min_len = min(len(df1_c), len(df2_c))

        result: Dict[str, Any] = {
            'missing_in_file2': max(0, len(df1_c) - len(df2_c)),
            'extra_in_file2': max(0, len(df2_c) - len(df1_c)),
            'rows_compared': min_len,
            'rows_with_differences': 0,   # filled below — NOT min_len
            'column_differences': {},
            'diff_rows': [],              # up to MAX_DIFF_ROWS_STORED cells
            'diff_rows_total': 0,         # real total before cap
            'sort_info': sort_info,       # información sobre el ordenamiento aplicado
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
                n1 = s1.apply(lambda v: self._normalize_value(v, case_sensitive, ignore_type_diff))
                n2 = s2.apply(lambda v: self._normalize_value(v, case_sensitive, ignore_type_diff))
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
        .button-green  { background: linear-gradient(135deg, #43a047 0%, #1b5e20 100%); }
        .button-danger { background: linear-gradient(135deg, #e53935 0%, #b71c1c 100%); }
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
        <button class="button button-danger" style="margin-top:18px;font-size:.9em" onclick="shutdownApp()">
            &#9209; Cerrar Aplicación
        </button>
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
                <label class="checkbox-label">
                    <input type="checkbox" id="autoSort" checked>
                    Ordenar datos automáticamente antes de comparar
                </label>
                <label class="checkbox-label">
                    <input type="checkbox" id="ignoreTypeDiff" checked>
                    Ignorar diferencias de tipo (comparar solo valores)
                </label>
            </div>
            
            <div id="sortColumnsSection" style="display:none; margin: 16px 0; padding: 14px; background: #f0f2ff; border-radius: 6px;">
                <p style="margin-bottom: 8px; font-weight: 600; color: #555;">
                    <span style="color: #667eea;">✓</span> Columnas sugeridas para ordenar:
                </p>
                <div id="suggestedColumns" style="margin-bottom: 12px; font-size: .9em; color: #666;"></div>
                <p style="margin: 8px 0; font-size: .9em; color: #666;">
                    Puedes personalizar las columnas de ordenamiento (separadas por comas):
                </p>
                <input type="text" id="sortColumnsInput" placeholder="Ejemplo: ID, Código, Fecha" 
                       style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: .9em;">
                <p style="margin-top: 6px; font-size: .85em; color: #888;">
                    💡 Esto permite comparar archivos con los mismos datos pero en diferente orden
                </p>
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

function shutdownApp() {
    if (!confirm('¿Cerrar la aplicación y el servidor?')) return;
    fetch('/shutdown', { method: 'POST' })
        .finally(() => {
            document.body.innerHTML =
                '<div style="display:flex;align-items:center;justify-content:center;height:100vh;' +
                'font-family:sans-serif;color:#555;font-size:1.3em;">' +
                'Aplicación cerrada. Puedes cerrar esta pestaña.</div>';
        });
}

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
    
    // Obtener columnas comunes si hay 2 archivos
    if (selectedFiles.length === 2) {
        fetchCommonColumns();
    } else {
        document.getElementById('sortColumnsSection').style.display = 'none';
    }
}

function fetchCommonColumns() {
    const fd = new FormData();
    fd.append('file1', selectedFiles[0]);
    fd.append('file2', selectedFiles[1]);
    
    fetch('/get-common-columns', { method:'POST', body:fd })
        .then(r => r.json())
        .then(d => {
            if (d.error) {
                console.error('Error obteniendo columnas:', d.error);
                return;
            }
            
            const section = document.getElementById('sortColumnsSection');
            const suggestedDiv = document.getElementById('suggestedColumns');
            const input = document.getElementById('sortColumnsInput');
            
            if (d.suggested_columns && d.suggested_columns.length > 0) {
                suggestedDiv.innerHTML = '<strong style="color: #667eea;">' + 
                    d.suggested_columns.join(', ') + '</strong>';
                input.value = d.suggested_columns.join(', ');
                section.style.display = 'block';
            } else {
                suggestedDiv.innerHTML = '<em>No se detectaron columnas automáticamente. Puedes especificarlas manualmente abajo.</em>';
                input.value = '';
                section.style.display = 'block';
            }
        })
        .catch(e => {
            console.error('Error:', e);
        });
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
    fd.append('auto_sort', document.getElementById('autoSort').checked ? '1' : '0');
    fd.append('ignore_type_diff', document.getElementById('ignoreTypeDiff').checked ? '1' : '0');
    
    const sortCols = document.getElementById('sortColumnsInput').value.trim();
    if (sortCols) {
        fd.append('sort_columns', sortCols);
    }
    
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
    
    const ignoreTypeNote = data.ignore_type_diff
        ? '<small style="color:#2e7d32">✓ Ignorando diferencias de tipo (comparando solo valores)</small>'
        : '';

    let html = `<div class="results">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px">
            <h3 style="color:#333">Comparación: <em>${esc(data.file1)}</em> vs <em>${esc(data.file2)}</em></h3>
            <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;">
                ${caseNote}
                ${ignoreTypeNote}
            </div>
        </div>`;

    // Sort info
    if (dataDiff.sort_info && dataDiff.sort_info.sorted) {
        const sortCols = dataDiff.sort_info.columns.join(', ');
        html += `<div class="success" style="margin-top: 12px;">
            ✓ Datos ordenados antes de comparar por: <strong>${esc(sortCols)}</strong>
            <br><small>Esto permite detectar los mismos datos independientemente del orden de las filas</small>
        </div>`;
    } else if (dataDiff.sort_info && dataDiff.sort_info.columns && dataDiff.sort_info.columns.length > 0) {
        html += `<div class="warning-box">
            ⚠ No se pudo ordenar por las columnas especificadas: ${esc(dataDiff.sort_info.columns.join(', '))}
            ${dataDiff.sort_info.error ? '<br><small>Error: ' + esc(dataDiff.sort_info.error) + '</small>' : ''}
        </div>`;
    }

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

_REPORT_CSS = '''
  *, *::before, *::after { box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
         margin: 0; background: #f0f2f5; color: #222; line-height: 1.5; }
  .page { max-width: 1150px; margin: 0 auto; padding: 30px 20px 60px; }
  h1 { color: #3949ab; font-size: 1.8em; margin: 0 0 4px; }
  .subtitle { color: #888; font-size: .9em; margin-bottom: 22px; }
  h2 { font-size: 1.15em; color: #3949ab; border-left: 4px solid #3949ab;
       padding: 7px 14px; margin: 34px 0 14px; background: #e8eaf6;
       border-radius: 0 6px 6px 0; }
  .card-row { display: flex; flex-wrap: wrap; gap: 12px; margin: 14px 0; }
  .card { background: white; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,.1);
          padding: 14px 20px; flex: 1; min-width: 140px; }
  .card .lbl { font-size: .75em; color: #888; text-transform: uppercase;
               letter-spacing: .05em; margin-bottom: 4px; }
  .card .val { font-size: 1.8em; font-weight: 700; color: #3949ab; }
  .card .val.red    { color: #c62828; }
  .card .val.green  { color: #2e7d32; }
  .card .val.orange { color: #e65100; }
  .card .sub { font-size: .78em; color: #777; margin-top: 3px; }
  .status-badge { display: inline-flex; align-items: center; gap: 8px;
                  padding: 8px 18px; border-radius: 20px; font-weight: 600;
                  font-size: .97em; margin-bottom: 14px; }
  .status-ok     { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; }
  .status-issues { background: #ffebee; color: #c62828; border: 1px solid #ef9a9a; }
  .diagnosis { list-style: none; margin: 0; padding: 0; }
  .diagnosis li { padding: 10px 16px; border-radius: 6px; margin-bottom: 8px;
                  font-size: .95em; line-height: 1.6; }
  .ok-item    { background: #e8f5e9; border-left: 4px solid #43a047; color: #1b5e20; }
  .issue-item { background: #fff8e1; border-left: 4px solid #f9a825; color: #4a3000; }
  .warn { background: #fff8e1; border-left: 4px solid #f9a825; padding: 10px 14px;
          border-radius: 4px; margin: 8px 0; color: #6d4c00; font-size: .9em; }
  .info { background: #e3f2fd; border-left: 4px solid #1976d2; padding: 10px 14px;
          border-radius: 4px; margin: 8px 0; color: #0d47a1; font-size: .9em; }
  .meta-box { background: white; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,.1);
              padding: 14px 20px; margin-bottom: 20px; }
  .meta-box table { box-shadow: none !important; border-radius: 0 !important;
                    margin: 0; width: auto; }
  .meta-box td { padding: 3px 16px 3px 0; border: none !important; background: transparent !important;
                 font-size: .92em; }
  .meta-box td:first-child { font-weight: 600; color: #555; white-space: nowrap; }
  .section-wrap { background: white; border-radius: 8px;
                  box-shadow: 0 1px 4px rgba(0,0,0,.1); padding: 20px; margin-bottom: 20px; }
  .col-lists { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
               gap: 14px; margin-top: 14px; }
  .col-group { border: 1px solid #e0e0e0; border-radius: 8px; padding: 14px; }
  .col-group h4 { margin: 0 0 10px; font-size: .82em; text-transform: uppercase;
                  letter-spacing: .04em; color: #555; }
  .pill { display: inline-block; border-radius: 4px; padding: 2px 9px; margin: 2px; font-size: .82em; }
  .pill-miss   { background: #ffebee; border: 1px solid #ef9a9a; color: #c62828; }
  .pill-extra  { background: #e8f5e9; border: 1px solid #a5d6a7; color: #2e7d32; }
  .pill-common { background: #e8eaf6; border: 1px solid #9fa8da; color: #3949ab; }
  table { width: 100%; border-collapse: collapse; background: white; font-size: .88em;
          box-shadow: 0 1px 4px rgba(0,0,0,.1); border-radius: 8px; overflow: hidden; }
  thead th { background: #3949ab; color: white; padding: 10px 14px; text-align: left;
             font-weight: 600; white-space: nowrap; }
  tbody tr:nth-child(even) { background: #fafafa; }
  tbody tr:hover { background: #f0f2ff; }
  td { padding: 8px 14px; border-bottom: 1px solid #eee; vertical-align: top; }
  tbody tr:last-child td { border-bottom: none; }
  td.row-num  { font-family: monospace; color: #888; font-size: .85em; white-space: nowrap; }
  td.val-old  { background: #fff3e0 !important; color: #bf360c; font-weight: 500; }
  td.val-new  { background: #e8f5e9 !important; color: #1b5e20; font-weight: 500; }
  td.val-null { color: #bbb; font-style: italic; }
  td.cnt      { font-weight: 700; color: #c62828; text-align: right; }
  .badge { display: inline-block; padding: 2px 9px; border-radius: 10px;
           font-size: .78em; font-weight: 600; white-space: nowrap; }
  .b-mod  { background: #fff3e0; color: #e65100; }
  .b-fill { background: #e8f5e9; color: #2e7d32; }
  .b-clr  { background: #ffebee; color: #c62828; }
  .trunc { background: #fff3e0; border-left: 4px solid #ff9800; padding: 10px 14px;
           border-radius: 4px; margin-bottom: 12px; font-size: .9em; color: #6d4c00; }
  @media print { body { background: white; } .page { padding: 0; } }
'''


def _rpt_structure(schema: dict, metrics: dict, file1: str, file2: str) -> str:
    only1  = schema.get('only_in_file1', [])
    only2  = schema.get('only_in_file2', [])
    common = schema.get('common_columns', [])
    row_diff = metrics.get('row_difference', 0)
    r1 = metrics.get('file1_rows', 0)
    r2 = metrics.get('file2_rows', 0)

    if row_diff < 0:
        diff_sub = f'Faltan {abs(row_diff)} fila(s) en Archivo 2'
        diff_cls = 'red'
    elif row_diff > 0:
        diff_sub = f'Sobran {row_diff} fila(s) en Archivo 2'
        diff_cls = 'orange'
    else:
        diff_sub = 'Mismo n&uacute;mero de filas'
        diff_cls = 'green'

    excl = len(only1) + len(only2)
    html = (
        f'<div class="section-wrap"><div class="card-row">'
        f'<div class="card"><div class="lbl">Filas en Archivo 1</div>'
        f'<div class="val">{r1:,}</div></div>'
        f'<div class="card"><div class="lbl">Filas en Archivo 2</div>'
        f'<div class="val">{r2:,}</div></div>'
        f'<div class="card"><div class="lbl">Diferencia de filas</div>'
        f'<div class="val {diff_cls}">{row_diff:+}</div>'
        f'<div class="sub">{diff_sub}</div></div>'
        f'<div class="card"><div class="lbl">Columnas comunes</div>'
        f'<div class="val">{len(common)}</div></div>'
        f'<div class="card"><div class="lbl">Columnas exclusivas</div>'
        f'<div class="val {"orange" if excl else "green"}">{excl}</div>'
        f'<div class="sub">{"Solo en un archivo" if excl else "Ninguna"}</div></div>'
        f'</div><div class="col-lists">'
    )
    if only1:
        pills = ''.join(f'<span class="pill pill-miss">&#10060; {c}</span>' for c in only1)
        html += (
            f'<div class="col-group"><h4>&#10060; Falta en Archivo 2 &mdash; {len(only1)} columna(s)</h4>'
            f'{pills}</div>'
        )
    if only2:
        pills = ''.join(f'<span class="pill pill-extra">&#10133; {c}</span>' for c in only2)
        html += (
            f'<div class="col-group"><h4>&#10133; Nueva en Archivo 2 &mdash; {len(only2)} columna(s)</h4>'
            f'{pills}</div>'
        )
    if common:
        pills = ''.join(f'<span class="pill pill-common">{c}</span>' for c in common[:60])
        more = (
            f'<span style="color:#888;font-size:.82em"> &hellip; y {len(common)-60} m&aacute;s</span>'
            if len(common) > 60 else ''
        )
        html += (
            f'<div class="col-group"><h4>&#10003; Columnas comunes &mdash; {len(common)}</h4>'
            f'{pills}{more}</div>'
        )
    html += '</div></div>'
    return html


def _rpt_col_summary(col_diffs: dict, rows_compared: int, diff_rows: list) -> str:
    type_counts: Dict[str, Dict[str, int]] = {}
    for r in diff_rows:
        col = r['col']
        if col not in type_counts:
            type_counts[col] = {'modified': 0, 'filled': 0, 'cleared': 0}
        if r.get('null1') and not r.get('null2'):
            type_counts[col]['filled'] += 1
        elif not r.get('null1') and r.get('null2'):
            type_counts[col]['cleared'] += 1
        else:
            type_counts[col]['modified'] += 1

    rows = ''
    for col, cnt in sorted(col_diffs.items(), key=lambda x: -x[1]):
        pct = f'{cnt / rows_compared * 100:.1f}%' if rows_compared else '&mdash;'
        tc  = type_counts.get(col, {})
        mod = tc.get('modified', 0)
        fil = tc.get('filled',   0)
        clr = tc.get('cleared',  0)
        rows += (
            f'<tr><td><strong>{col}</strong></td>'
            f'<td class="cnt">{cnt}</td><td>{pct}</td>'
            f'<td>{"<span class=badge b-mod>" + str(mod) + " valor(es)</span>" if mod else "&mdash;"}</td>'
            f'<td>{"<span class=badge b-fill>" + str(fil) + " celda(s)</span>" if fil else "&mdash;"}</td>'
            f'<td>{"<span class=badge b-clr>" + str(clr) + " celda(s)</span>" if clr else "&mdash;"}</td>'
            f'</tr>'
        )
    return (
        '<div class="section-wrap"><div style="overflow-x:auto"><table>'
        '<thead><tr>'
        '<th>Columna</th><th>Celdas distintas</th><th>% comparadas</th>'
        '<th>Valor modificado</th><th>Vac&iacute;o &rarr; dato</th><th>Dato &rarr; vac&iacute;o</th>'
        '</tr></thead>'
        f'<tbody>{rows}</tbody></table></div></div>'
    )


def _rpt_detail_table(diff_rows: list, total_cells: int, file1: str, file2: str) -> str:
    trunc = ''
    if total_cells > len(diff_rows):
        trunc = (
            f'<div class="trunc">&#9888; Se muestran <strong>{len(diff_rows):,}</strong> '
            f'de <strong>{total_cells:,}</strong> diferencias totales.</div>'
        )
    rows = ''
    for r in diff_rows:
        null1, null2 = r.get('null1'), r.get('null2')
        v1, v2 = r['val1'], r['val2']
        if null1 and not null2:
            badge = '<span class="badge b-fill">Vac&iacute;o &rarr; dato</span>'
            c1, c2 = 'val-null', 'val-new'
        elif not null1 and null2:
            badge = '<span class="badge b-clr">Dato &rarr; vac&iacute;o</span>'
            c1, c2 = 'val-old', 'val-null'
        else:
            badge = '<span class="badge b-mod">Valor modificado</span>'
            c1, c2 = 'val-old', 'val-new'
        rows += (
            f'<tr>'
            f'<td class="row-num">Fila&nbsp;{r["row_excel"]}</td>'
            f'<td><strong>{r["col"]}</strong></td>'
            f'<td>{badge}</td>'
            f'<td class="{c1}">{v1}</td>'
            f'<td class="{c2}">{v2}</td>'
            f'</tr>'
        )
    return (
        f'<div class="section-wrap">{trunc}'
        f'<div style="overflow-x:auto"><table>'
        f'<thead><tr>'
        f'<th>Fila (Excel)</th><th>Columna</th><th>Tipo de cambio</th>'
        f'<th>{file1}<br><small style="font-weight:400;opacity:.85">(valor anterior)</small></th>'
        f'<th>{file2}<br><small style="font-weight:400;opacity:.85">(valor nuevo)</small></th>'
        f'</tr></thead>'
        f'<tbody>{rows}</tbody></table></div></div>'
    )


def _build_download_report(data: dict) -> str:
    schema      = data.get('schema_differences', {})
    dd          = data.get('data_differences', {})
    metrics     = data.get('metrics', {})
    diff_rows   = dd.get('diff_rows', [])
    col_diffs   = dd.get('column_differences', {})
    file1       = data.get('file1', 'Archivo 1')
    file2       = data.get('file2', 'Archivo 2')
    total_cells = dd.get('diff_rows_total', len(diff_rows))
    rows_cmp    = dd.get('rows_compared', 0)
    row_diff    = metrics.get('row_difference', 0)
    only1       = schema.get('only_in_file1', [])
    only2       = schema.get('only_in_file2', [])
    rows_w_diff = dd.get('rows_with_differences', 0)

    # ── Diagnóstico en lenguaje natural ─────────────────────────────────
    issues: List[str] = []
    if row_diff < 0:
        issues.append(
            f'<li class="issue-item">&#10060; <strong>Faltan {abs(row_diff)} fila(s)</strong> en Archivo 2 '
            f'(Archivo&nbsp;1:&nbsp;{metrics.get("file1_rows",0):,}&nbsp;filas &mdash; '
            f'Archivo&nbsp;2:&nbsp;{metrics.get("file2_rows",0):,}&nbsp;filas).</li>'
        )
    elif row_diff > 0:
        issues.append(
            f'<li class="issue-item">&#9888; <strong>Sobran {row_diff} fila(s)</strong> en Archivo 2 '
            f'(Archivo&nbsp;1:&nbsp;{metrics.get("file1_rows",0):,}&nbsp;filas &mdash; '
            f'Archivo&nbsp;2:&nbsp;{metrics.get("file2_rows",0):,}&nbsp;filas). '
            f'Las filas extra no fueron comparadas.</li>'
        )
    for col in only1:
        issues.append(
            f'<li class="issue-item">&#10060; <strong>Columna faltante en Archivo 2:</strong> '
            f'&laquo;{col}&raquo; &mdash; existe en Archivo 1 pero <strong>no aparece</strong> en Archivo 2.</li>'
        )
    for col in only2:
        issues.append(
            f'<li class="issue-item">&#10133; <strong>Columna nueva en Archivo 2:</strong> '
            f'&laquo;{col}&raquo; &mdash; existe en Archivo 2 pero <strong>no aparece</strong> en Archivo 1.</li>'
        )
    if total_cells:
        issues.append(
            f'<li class="issue-item">&#128260; '
            f'<strong>{total_cells:,} celda(s) con valor distinto</strong> en '
            f'<strong>{rows_w_diff:,} fila(s)</strong> y '
            f'<strong>{len(col_diffs)} columna(s)</strong>.</li>'
        )

    has_issues = bool(issues)
    if not has_issues:
        issues.append(
            '<li class="ok-item">&#10003; Los archivos son <strong>id&eacute;nticos</strong> '
            'en estructura y datos.</li>'
        )

    diagnosis_html = '<ul class="diagnosis">' + ''.join(issues) + '</ul>'
    status_cls  = 'status-issues' if has_issues else 'status-ok'
    status_text = (
        f'&#9888;&nbsp;{len(issues)} hallazgo(s) encontrado(s)' if has_issues
        else '&#10003;&nbsp;Sin diferencias'
    )

    warnings_html = ''.join(f'<div class="warn">&#9888; {w}</div>' for w in data.get('warnings', []))
    sort_info = dd.get('sort_info', {})
    sort_html = ''
    if sort_info.get('sorted'):
        sort_html = (
            f'<div class="info">&#128260; Datos ordenados antes de comparar por: '
            f'<strong>{", ".join(sort_info["columns"])}</strong></div>'
        )
    ts = data.get('timestamp', '?')[:19].replace('T', ' ')

    structure_html   = _rpt_structure(schema, metrics, file1, file2)
    col_summary_html = _rpt_col_summary(col_diffs, rows_cmp, diff_rows) if col_diffs else ''
    detail_html      = _rpt_detail_table(diff_rows, total_cells, file1, file2) if diff_rows else ''
    no_diff_html     = (
        '<div class="section-wrap"><p style="color:#2e7d32;font-weight:600">'
        '&#10003; Sin diferencias en los datos de columnas comunes.</p></div>'
        if not diff_rows and not only1 and not only2 and not row_diff else ''
    )

    return f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Informe &mdash; {file1} vs {file2}</title>
<style>{_REPORT_CSS}</style>
</head>
<body>
<div class="page">

<h1>&#128202; Informe de Comparaci&oacute;n Excel</h1>
<p class="subtitle">Generado el {ts}</p>

<div class="meta-box">
  <table>
    <tr><td>Archivo 1:</td><td>{file1}</td></tr>
    <tr><td>Archivo 2:</td><td>{file2}</td></tr>
    <tr><td>Filas comparadas:</td><td>{rows_cmp:,}</td></tr>
    <tr><td>Comparaci&oacute;n:</td><td>{"Sensible a may&uacute;sculas" if data.get("case_sensitive") else "Insensible a may&uacute;sculas"}</td></tr>
  </table>
</div>

{warnings_html}{sort_html}

<h2>&#128203; Diagn&oacute;stico General</h2>
<div class="section-wrap">
  <div class="{status_cls} status-badge">{status_text}</div>
  {diagnosis_html}
</div>

<h2>&#128208; Estructura de Archivos</h2>
{structure_html}

{"<h2>&#128202; Resumen de Cambios por Columna</h2>" + col_summary_html if col_summary_html else ""}

{"<h2>&#128269; Detalle Celda por Celda</h2>" + detail_html if detail_html else ""}

{no_diff_html}

</div>
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

    secure_name = secure_filename(f.filename) or 'file.xlsx'
    unique_name = f"{uuid.uuid4().hex}_{secure_name}"
    path = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], unique_name))
    
    f.save(path)
    try:
        # Verificar que el archivo existe
        if not os.path.exists(path):
            return jsonify({'error': 'Error al guardar el archivo temporal'}), 500
        return jsonify(analyzer.analyze_single_file(path, display_name=f.filename or secure_name))
    finally:
        if os.path.exists(path):
            os.remove(path)


@app.route('/get-common-columns', methods=['POST'])
def get_common_columns():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': 'Se requieren 2 archivos'}), 400

    f1 = request.files['file1']
    f2 = request.files['file2']

    req_id = uuid.uuid4().hex
    secure_name1 = secure_filename(f1.filename) or 'file1.xlsx'
    secure_name2 = secure_filename(f2.filename) or 'file2.xlsx'
    p1 = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], f"{req_id}_1_{secure_name1}"))
    p2 = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], f"{req_id}_2_{secure_name2}"))

    f1.save(p1)
    f2.save(p2)
    try:
        if not os.path.exists(p1) or not os.path.exists(p2):
            return jsonify({'error': 'Error al guardar los archivos temporales'}), 500

        df1 = analyzer._smart_read_excel(p1)
        df2 = analyzer._smart_read_excel(p2)

        df1 = analyzer._preprocess_dataframe(df1)
        df2 = analyzer._preprocess_dataframe(df2)
        
        common_cols = sorted(set(df1.columns) & set(df2.columns))
        suggested_cols = analyzer._detect_sort_columns(df1, df2)
        return jsonify({
            'common_columns': common_cols,
            'suggested_columns': suggested_cols
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        import time
        import gc
        gc.collect()
        time.sleep(0.1)
        for p in (p1, p2):
            if os.path.exists(p):
                max_attempts = 3
                for attempt in range(max_attempts):
                    try:
                        os.remove(p)
                        break
                    except PermissionError:
                        if attempt < max_attempts - 1:
                            time.sleep(0.5)


@app.route('/compare', methods=['POST'])
def compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': 'Se requieren 2 archivos'}), 400

    f1 = request.files['file1']
    f2 = request.files['file2']
    case_sensitive = request.form.get('case_sensitive', '1') != '0'
    auto_sort = request.form.get('auto_sort', '1') != '0'
    ignore_type_diff = request.form.get('ignore_type_diff', '0') != '0'
    
    # Obtener columnas de ordenamiento si se especificaron
    sort_columns_str = request.form.get('sort_columns', '')
    sort_columns = [col.strip() for col in sort_columns_str.split(',') if col.strip()] if sort_columns_str else None

    req_id = uuid.uuid4().hex
    secure_name1 = secure_filename(f1.filename) or 'file1.xlsx'
    secure_name2 = secure_filename(f2.filename) or 'file2.xlsx'
    p1 = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], f"{req_id}_1_{secure_name1}"))
    p2 = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], f"{req_id}_2_{secure_name2}"))

    try:
        f1.save(p1)
        f2.save(p2)
        
        # Verificar que los archivos existen antes de procesarlos
        if not os.path.exists(p1):
            return jsonify({'error': f'No se pudo guardar el archivo: {f1.filename}'}), 500
        if not os.path.exists(p2):
            return jsonify({'error': f'No se pudo guardar el archivo: {f2.filename}'}), 500
        
        result = analyzer.compare_files(
            p1, p2,
            case_sensitive=case_sensitive,
            sort_columns=sort_columns,
            auto_sort=auto_sort,
            ignore_type_diff=ignore_type_diff,
            file1_name=f1.filename or secure_name1,
            file2_name=f2.filename or secure_name2,
        )
        return jsonify(result)
    except Exception as e:
        import traceback
        error_msg = f'{str(e)}\n\nRuta archivo 1: {p1}\nRuta archivo 2: {p2}'
        return jsonify({'error': error_msg, 'traceback': traceback.format_exc()}), 500
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

    # Asegurar que la base de datos y tabla existen
    analyzer._ensure_db()
    
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


@app.route('/shutdown', methods=['POST'])
def shutdown():
    def _stop():
        import time as _t
        _t.sleep(0.6)
        os._exit(0)
    threading.Thread(target=_stop, daemon=True).start()
    return jsonify({'message': 'Cerrando aplicación...'})


if __name__ == '__main__':
    threading.Timer(1.5, lambda: webbrowser.open('http://127.0.0.1:5000')).start()
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
