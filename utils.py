"""
Utilidades adicionales para Excel Analyzer
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)

class ExcelUtils:
    """Utilidades para manejo de archivos Excel"""
    
    @staticmethod
    def get_file_size(file_path: str) -> Dict[str, Any]:
        """Obtiene info de tamaño del archivo"""
        import os
        size_bytes = os.path.getsize(file_path)
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024:
                return {'size': size_bytes, 'unit': unit, 'formatted': f"{size_bytes:.2f} {unit}"}
            size_bytes /= 1024
        
        return {'size': size_bytes, 'unit': 'GB', 'formatted': f"{size_bytes:.2f} GB"}
    
    @staticmethod
    def detect_data_quality_issues(df: pd.DataFrame) -> Dict[str, Any]:
        """Detecta problemas de calidad en los datos"""
        issues = {
            'missing_values': {},
            'duplicates': 0,
            'outliers': {},
            'inconsistent_types': []
        }
        
        # Valores faltantes
        for col in df.columns:
            missing = df[col].isna().sum()
            if missing > 0:
                issues['missing_values'][col] = {
                    'count': int(missing),
                    'percentage': float(missing / len(df) * 100)
                }
        
        # Duplicados
        issues['duplicates'] = int(df.duplicated().sum())
        
        # Outliers en columnas numéricas
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            Q1 = df[col].quantile(0.25)
            Q3 = df[col].quantile(0.75)
            IQR = Q3 - Q1
            outliers = ((df[col] < (Q1 - 1.5 * IQR)) | 
                       (df[col] > (Q3 + 1.5 * IQR))).sum()
            if outliers > 0:
                issues['outliers'][col] = int(outliers)
        
        return issues
    
    @staticmethod
    def get_column_correlation(df: pd.DataFrame, threshold: float = 0.8) -> Dict[str, List[str]]:
        """Detecta columnas altamente correlacionadas"""
        numeric_df = df.select_dtypes(include=[np.number])
        if numeric_df.shape[1] < 2:
            return {}
        
        corr_matrix = numeric_df.corr().abs()
        
        high_corr = {}
        for i in range(len(corr_matrix.columns)):
            for j in range(i + 1, len(corr_matrix.columns)):
                if corr_matrix.iloc[i, j] > threshold:
                    col1 = corr_matrix.columns[i]
                    col2 = corr_matrix.columns[j]
                    corr_value = float(corr_matrix.iloc[i, j])
                    
                    key = f"{col1} - {col2}"
                    high_corr[key] = corr_value
        
        return high_corr
    
    @staticmethod
    def generate_data_profile(df: pd.DataFrame) -> Dict[str, Any]:
        """Genera un perfil completo de los datos"""
        profile = {
            'shape': df.shape,
            'memory_mb': float(df.memory_usage(deep=True).sum() / 1024 / 1024),
            'columns_info': {},
            'quality_issues': ExcelUtils.detect_data_quality_issues(df),
            'correlations': ExcelUtils.get_column_correlation(df)
        }
        
        for col in df.columns:
            col_info = {
                'dtype': str(df[col].dtype),
                'non_null': int(df[col].notna().sum()),
                'null_count': int(df[col].isna().sum()),
                'unique': int(df[col].nunique())
            }
            
            if pd.api.types.is_numeric_dtype(df[col]):
                col_info.update({
                    'min': float(df[col].min()),
                    'max': float(df[col].max()),
                    'mean': float(df[col].mean()),
                    'std': float(df[col].std())
                })
            elif pd.api.types.is_object_dtype(df[col]):
                col_info['top_value'] = str(df[col].value_counts().index[0]) if len(df[col].value_counts()) > 0 else None
            
            profile['columns_info'][col] = col_info
        
        return profile

class ComparisonReporter:
    """Generador de reportes de comparación"""
    
    @staticmethod
    def generate_html_report(comparison_result: Dict[str, Any], output_file: str = 'comparison_report.html'):
        """Genera un reporte HTML de la comparación"""
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Reporte de Comparación</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
                .header {{ background: #667eea; color: white; padding: 20px; border-radius: 5px; }}
                .section {{ background: white; margin: 20px 0; padding: 15px; border-radius: 5px; }}
                table {{ width: 100%; border-collapse: collapse; }}
                th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
                th {{ background: #667eea; color: white; }}
                .metric {{ display: inline-block; width: 20%; margin: 10px; text-align: center; }}
                .metric-value {{ font-size: 24px; font-weight: bold; color: #667eea; }}
                .metric-label {{ color: #666; margin-top: 5px; }}
                .difference {{ background: #ffe0e0; }}
                .ok {{ background: #e0ffe0; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Reporte de Comparación</h1>
                <p>Comparación entre archivos Excel</p>
            </div>
            
            <div class="section">
                <h2>Archivos Comparados</h2>
                <p><strong>Archivo 1:</strong> {comparison_result.get('file1', 'N/A')}</p>
                <p><strong>Archivo 2:</strong> {comparison_result.get('file2', 'N/A')}</p>
                <p><strong>Fecha:</strong> {comparison_result.get('timestamp', 'N/A')}</p>
            </div>
            
            <div class="section">
                <h2>Métricas Principales</h2>
                <div class="metric">
                    <div class="metric-value">{comparison_result.get('metrics', {}).get('row_difference', 0):+}</div>
                    <div class="metric-label">Diferencia Filas</div>
                </div>
                <div class="metric">
                    <div class="metric-value">{comparison_result.get('metrics', {}).get('col_difference', 0):+}</div>
                    <div class="metric-label">Diferencia Columnas</div>
                </div>
                <div class="metric">
                    <div class="metric-value">{comparison_result.get('schema_differences', {}).get('total_different', 0)}</div>
                    <div class="metric-label">Columnas Diferentes</div>
                </div>
            </div>
        </body>
        </html>
        """
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_file

def setup_logging(log_file: str = 'analyzer.log'):
    """Configura logging"""
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
