# Sales_Insights_Pro.py
#
# A professional, multi-lingual, multi-file-type Sales Dashboard and Forecasting tool.
# Version 2.0: Added caching for performance and improved type hinting.
#
# Author: Sameh Sobhy Attia (Original)
# Refactored by: Gemini (Professional Upgrade)
#
# ---Dependencies---
# To run this app, you need Streamlit and other data libraries.
# Install them using pip:
# pip install streamlit pandas numpy plotly openpyxl reportlab lxml pdfplumber
#
# ---To Run---
# Save this file as "Sales_Insights_Pro.py"
# In your terminal, run:
# streamlit run Sales_Insights_Pro.py
#
# ---Features---
# - Caching for high-performance data processing.
# - Supports Excel, CSV, PDF, and HTML (table extraction) file uploads.
# - Fully bilingual (English/Arabic) UI.
# - Robust session state management (data persists across interactions).
# - Clean, tabbed interface for:
#   1. KPIs & Statistics
#   2. Pivot Tables
#   3. Charting
#   4. Forecasting
#   5. Automated Insights
#   6. Report Exports (Excel, HTML, PDF)
# - Dark mode support.

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import pdfplumber  # For reading PDF tables
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from typing import List, Dict, Tuple, Optional, Any, BinaryIO
from lxml import etree # Used for HTML parsing, openpyxl needs it

# ================================================
# 1. APP CONFIGURATION & INITIALIZATION
# ================================================

st.set_page_config(page_title='Sales Insights Pro', layout='wide')

# Initialize session state
if 'lang' not in st.session_state:
    st.session_state['lang'] = 'en'
if 'df' not in st.session_state:
    st.session_state['df'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = None

# ================================================
# 2. TRANSLATIONS & LANGUAGE HELPER
# ================================================

TRANSLATIONS = {
    'en': {
        'title': 'Sales Insights & Forecasting Pro',
        'upload': 'Upload Data (Excel, CSV, PDF, HTML)',
        'upload_prompt': 'Upload a file to get started. Supported formats: Excel, CSV, PDF, HTML (tables).',
        'load_sample': 'Load Sample Data',
        'data_loaded': 'Successfully loaded',
        'rows': 'rows',
        'cols': 'columns',
        'total_everything': 'Total of all Numeric Columns',
        'grand_total': 'Grand Total',
        'kpi_selection': 'Select KPI Columns (for totals, forecasting)',
        'date_column': 'Select Date Column (for time series)',
        'pivot_config': 'Pivot Table Configuration',
        'row_field': 'Row Field(s)',
        'col_field': 'Column Field(s)',
        'agg_type': 'Aggregation Type',
        'value_col': 'Value Column',
        'generate_pivot': 'Generate Pivot Table',
        'stats_summary': 'Statistics Summary',
        'charts': 'Charts & Visuals',
        'chart_type': 'Chart Type',
        'x_axis': 'X-Axis',
        'y_axis': 'Y-Axis (multi-select)',
        'plot': 'Plot Chart',
        'forecasting': 'Simple Forecasting (Trend)',
        'forecast_column': 'Select numeric column to forecast',
        'forecast_periods': 'Forecast Periods (steps)',
        'run_forecast': 'Run Forecast',
        'insights': 'Automated Insights',
        'missing_values': 'Missing Values by Column',
        'correlations': 'Correlation Matrix (Numeric)',
        'download_excel': 'Download Summary as Excel',
        'download_html': 'Download Report as HTML',
        'download_pdf': 'Download Report as PDF',
        'language': 'Language',
        'theme': 'Dark Mode',
        'show_data': 'Show Raw Data',
        'download_pivot': 'Download Pivot as Excel',
        'config': 'Column Configuration',
        'kpi_tab': 'KPIs & Stats',
        'pivot_tab': 'Pivot Table',
        'charts_tab': 'Charts',
        'forecast_tab': 'Forecasting',
        'insights_tab': 'Data Insights',
        'export_tab': 'Export Report',
        'selected_kpis': 'Totals for Selected KPIs',
        'no_kpis_selected': 'No KPI columns selected.',
        'no_numeric_stats': 'No numeric columns for statistics.',
        'plot_warn': 'Please select at least one Y-Axis column.',
        'forecast_warn': 'Please select a numeric column to forecast.',
        'forecast_no_date': 'No date column selected. Forecasting on data index.',
        'forecast_no_data': 'Not enough data to forecast (need at least 3 data points).',
        'forecast_fail': 'Forecasting failed',
        'forecast_table': 'Forecast Table',
        'actual': 'Actual',
        'forecast': 'Forecast',
        'confidence': 'Confidence Interval',
        'no_corr': 'Not enough numeric columns for correlation.',
        'file_error': 'Could not read file. Please ensure it is a valid format.',
        'pdf_warn': 'PDF parsing found 0 tables. Please check the file.',
        'html_warn': 'HTML parsing found 0 tables. Please check the file.',
        'footer_credit': 'Created by',
    },
    'ar': {
        'title': 'تحليلات المبيعات والتنبؤ الاحترافي',
        'upload': 'رفع البيانات (Excel, CSV, PDF, HTML)',
        'upload_prompt': 'ارفع ملفاً للبدء. الصيغ المدعومة: Excel, CSV, PDF, HTML (جداول).',
        'load_sample': 'تحميل بيانات عينة',
        'data_loaded': 'تم تحميل',
        'rows': 'صفوف',
        'cols': 'أعمدة',
        'total_everything': 'مجموع كل الأعمدة الرقمية',
        'grand_total': 'المجموع الكلي',
        'kpi_selection': 'اختر أعمدة المؤشرات (للإجماليات والتنبؤ)',
        'date_column': 'اختر عمود التاريخ (للسلاسل الزمنية)',
        'pivot_config': 'إعداد الجدول المحوري',
        'row_field': 'حقل (حقول) الصف',
        'col_field': 'حقل (حقول) العمود',
        'agg_type': 'نوع التجميع',
        'value_col': 'عمود القيمة',
        'generate_pivot': 'إنشاء جدول محوري',
        'stats_summary': 'ملخص الإحصائيات',
        'charts': 'المخططات والمرئيات',
        'chart_type': 'نوع المخطط',
        'x_axis': 'المحور السيني',
        'y_axis': 'المحور الصادي (اختيار متعدد)',
        'plot': 'ارسم المخطط',
        'forecasting': 'التنبؤ البسيط (الاتجاه)',
        'forecast_column': 'اختر العمود الرقمي للتنبؤ',
        'forecast_periods': 'فترات التنبؤ (خطوات)',
        'run_forecast': 'تشغيل التنبؤ',
        'insights': 'رؤى تلقائية',
        'missing_values': 'القيم المفقودة حسب العمود',
        'correlations': 'مصفوفة الارتباط (رقمي)',
        'download_excel': 'تحميل الملخص كملف Excel',
        'download_html': 'تحميل التقرير كملف HTML',
        'download_pdf': 'تحميل التقرير كملف PDF',
        'language': 'اللغة',
        'theme': 'الوضع الداكن',
        'show_data': 'عرض البيانات الخام',
        'download_pivot': 'تحميل الجدول المحوري كـ Excel',
        'config': 'تكوين الأعمدة',
        'kpi_tab': 'المؤشرات والإحصائيات',
        'pivot_tab': 'الجدول المحوري',
        'charts_tab': 'المخططات',
        'forecast_tab': 'التنبؤ',
        'insights_tab': 'تحليلات البيانات',
        'export_tab': 'تصدير التقرير',
        'selected_kpis': 'إجماليات المؤشرات المحددة',
        'no_kpis_selected': 'لم يتم تحديد أعمدة مؤشرات.',
        'no_numeric_stats': 'لا توجد أعمدة رقمية للإحصاءات.',
        'plot_warn': 'يرجى اختيار عمود واحد على الأقل للمحور الصادي.',
        'forecast_warn': 'يرجى اختيار عمود رقمي للتنبؤ.',
        'forecast_no_date': 'لم يتم تحديد عمود تاريخ. سيتم التنبؤ بناءً على تسلسل البيانات.',
        'forecast_no_data': 'لا توجد بيانات كافية للتنبؤ (تحتاج 3 نقاط بيانات على الأقل).',
        'forecast_fail': 'فشل التنبؤ',
        'forecast_table': 'جدول التنبؤ',
        'actual': 'الفعلي',
        'forecast': 'التنبؤ',
        'confidence': 'نطاق الثقة',
        'no_corr': 'لا توجد أعمدة رقمية كافية للارتباط.',
        'file_error': 'لا يمكن قراءة الملف. يرجى التأكد من أن الصيغة صحيحة.',
        'pdf_warn': 'لم يتم العثور على جداول في ملف PDF. يرجى فحص الملف.',
        'html_warn': 'لم يتم العثور على جداول في ملف HTML. يرجى فحص الملف.',
        'footer_credit': 'إعداد',
    }
}

def t(key: str) -> str:
    """
    Translation helper function.
    Fetches a translation string based on the current language in session state.
    """
    lang = st.session_state.get('lang', 'en')
    return TRANSLATIONS.get(lang, TRANSLATIONS['en']).get(key, key)

# ================================================
# 3. DATA LOADING & PARSING HELPERS (WITH CACHING)
# ================================================

@st.cache_data
def parse_pdf(file_content: bytes) -> Optional[pd.DataFrame]:
    """Extract tables from a PDF file."""
    all_tables = []
    try:
        with io.BytesIO(file_content) as f:
            with pdfplumber.open(f) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            all_tables.append(pd.DataFrame(table[1:], columns=table[0]))
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None
    
    if not all_tables:
        st.warning(t('pdf_warn'))
        return None
    
    df = pd.concat(all_tables, ignore_index=True)
    return df

@st.cache_data
def parse_html(file_content: bytes) -> Optional[pd.DataFrame]:
    """Extract tables from an HTML file."""
    try:
        tables = pd.read_html(io.BytesIO(file_content), encoding='utf-8')
        if not tables:
            st.warning(t('html_warn'))
            return None
        
        df = pd.concat(tables, ignore_index=True)
        return df
    except Exception as e:
        st.error(f"{t('file_error')}: {e}")
        return None

@st.cache_data
def parse_excel_csv(file_content: bytes, file_name: str) -> Optional[pd.DataFrame]:
    """Read and clean Excel/CSV files with smart header detection."""
    name = file_name.lower()
    df = None
    file_like_object = io.BytesIO(file_content)
    
    try:
        if name.endswith('.csv'):
            df = pd.read_csv(file_like_object, header=None, encoding='utf-8', engine='python')
        else:
            df = pd.read_excel(file_like_object, header=None, engine='openpyxl')
    except Exception as e:
        st.error(f"{t('file_error')}: {e}")
        return None

    # Drop completely empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    if df.empty:
        return None

    # Detect header row: pick the row with the most non-null values
    header_row = df.notna().sum(axis=1).idxmax()
    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    # Clean column names: replace Unnamed or blanks
    df.columns = [
        col if (isinstance(col, str) and col.strip() != "" and not col.strip().startswith("Unnamed"))
        else f"Column_{i}"
        for i, col in enumerate(df.columns)
    ]

    df = df.dropna(how="all").reset_index(drop=True)

    # Try converting numeric columns
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='ignore')

    # Drop duplicated columns
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def load_data(uploaded_file: BinaryIO):
    """
    Master function to load data from any supported file type.
    This function handles the file I/O and session state logic,
    while calling cached functions for the actual parsing.
    """
    if uploaded_file is None:
        return

    name = uploaded_file.name
    file_content = uploaded_file.getvalue()
    df = None

    try:
        if name.lower().endswith('.pdf'):
            df = parse_pdf(file_content)
        elif name.lower().endswith(('.html', '.htm')):
            df = parse_html(file_content)
        elif name.lower().endswith(('.csv', '.xls', '.xlsx')):
            df = parse_excel_csv(file_content, name)
        else:
            st.error(f"Unsupported file type: {name}")
            return

        if df is not None and not df.empty:
            # Post-processing for all loaded data
            df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
            for c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='ignore')
            
            st.session_state['df'] = df
            st.session_state['file_name'] = uploaded_file.name
            st.success(f"{t('data_loaded')} '{uploaded_file.name}' ({df.shape[0]} {t('rows')}, {df.shape[1]} {t('cols')})")
        elif df is None:
             # Error was already shown by the parsing function
             st.session_state['df'] = None
             st.session_state['file_name'] = None
        elif df is not None and df.empty:
             # Warning was already shown by the parsing function
             st.session_state['df'] = None
             st.session_state['file_name'] = None

    except Exception as e:
        st.error(f"{t('file_error')}: {e}")
        st.session_state['df'] = None
        st.session_state['file_name'] = None

@st.cache_data
def get_sample_data() -> pd.DataFrame:
    """Generates sample data."""
    df = pd.DataFrame({
        'Date': pd.date_range(end=pd.Timestamp.today(), periods=24, freq='MS'),
        'Category': ['A', 'B', 'C'] * 8,
        'Branch': ['North', 'South'] * 12,
        'Sales': np.random.randint(100, 1000, 24),
        'Quantity': np.random.randint(1, 50, 24),
        'Profit': np.random.randint(-50, 300, 24)
    })
    return df

def load_sample_data():
    """Loads sample data into session state."""
    df = get_sample_data()
    st.session_state['df'] = df
    st.session_state['file_name'] = 'Sample_Data.csv'
    st.success(f"{t('data_loaded')} 'Sample_Data.csv' ({df.shape[0]} {t('rows')}, {df.shape[1]} {t('cols')})")

# ================================================
# 4. ANALYSIS & PLOTTING HELPERS (WITH CACHING)
# ================================================

@st.cache_data
def grand_totals(df: pd.DataFrame) -> Tuple[Dict[str, float], float]:
    """Calculates totals for all numeric columns."""
    numeric = df.select_dtypes(include=[np.number])
    totals = numeric.sum(numeric_only=True)
    grand = totals.sum()
    return totals.to_dict(), grand

@st.cache_data
def stats_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Generates descriptive statistics."""
    numeric = df.select_dtypes(include=[np.number])
    if numeric.empty:
        return pd.DataFrame()
    summary = numeric.agg(['count', 'mean', 'median', 'max', 'min', 'std']).transpose()
    return summary

@st.cache_data
def generate_pivot(df: pd.DataFrame, rows: List[str], cols: List[str], values: Optional[str], aggfunc: str) -> Optional[pd.DataFrame]:
    """Generates a pivot table."""
    agg_map = {
        'sum': np.sum, 'mean': np.mean, 'median': np.median,
        'count': 'count', 'min': np.min, 'max': np.max, 'std': np.std,
    }
    func = agg_map.get(aggfunc, np.sum)
    try:
        pvt = pd.pivot_table(df, index=rows if rows else None, 
                             columns=cols if cols else None,
                             values=values if values else None, 
                             aggfunc=func, margins=True, fill_value=0)
        return pvt
    except Exception as e:
        st.error(f"Pivot error: {e}")
        return None

def run_forecast(df: pd.DataFrame, date_col: Optional[str], fc_col: str, fc_periods: int):
    """
    Runs and plots a simple polynomial forecast.
    Not cached as it's a quick calculation and should respond to UI changes.
    """
    if not fc_col:
        st.warning(t('forecast_warn'))
        return

    try:
        if date_col:
            # --- Forecasting with a Date Column ---
            tmp = df[[date_col, fc_col]].copy()
            tmp[date_col] = pd.to_datetime(tmp[date_col], errors='coerce')
            tmp = tmp.dropna(subset=[date_col, fc_col])
            tmp = tmp.groupby(date_col, as_index=False)[fc_col].mean().sort_values(date_col)
            tmp_series = tmp.set_index(date_col)[fc_col]
            tmp_series = tmp_series[~tmp_series.index.duplicated(keep='first')]
            
            if tmp_series.shape[0] < 3:
                st.warning(t('forecast_no_data'))
                return

            n = tmp_series.shape[0]
            deg = 1 if n < 6 else 2
            x = np.arange(n)
            coeffs = np.polyfit(x, tmp_series.values, deg)
            model = np.poly1d(coeffs)

            fitted = model(x)
            resid = tmp_series.values - fitted
            resid_std = np.nanstd(resid)
            ci = 1.96 * resid_std

            try:
                freq = pd.infer_freq(tmp_series.index)
                if freq is None: freq = 'D'
            except Exception:
                freq = 'D'
            
            last_date = tmp_series.index.max()
            future_index = pd.date_range(start=last_date, periods=int(fc_periods) + 1, freq=freq)[1:]

            future_x = np.arange(n, n + int(fc_periods))
            preds = model(future_x)
            
            forecast_df = pd.DataFrame({
                date_col: future_index,
                'forecast': preds,
                'lower_band': preds - ci,
                'upper_band': preds + ci
            })
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=tmp_series.index, y=tmp_series.values,
                                     mode='lines', name=t('actual'), line=dict(color='blue')))
            fig.add_trace(go.Scatter(x=forecast_df[date_col], y=forecast_df['forecast'],
                                     mode='lines', name=t('forecast'), line=dict(dash='dash', color='red', width=3)))
            fig.add_trace(go.Scatter(
                x=list(forecast_df[date_col]) + list(forecast_df[date_col][::-1]),
                y=list(forecast_df['upper_band']) + list(forecast_df['lower_band'][::-1]),
                fill='toself', fillcolor='rgba(255,0,0,0.15)',
                line=dict(color='rgba(255,255,255,0)'),
                hoverinfo="skip", showlegend=True, name=t('confidence')
            ))
            fig.update_layout(title=f"{fc_col} - {t('forecast')}", xaxis_title=date_col, yaxis_title=fc_col)
            st.plotly_chart(fig, use_container_width=True)
            st.subheader(t('forecast_table'))
            st.dataframe(forecast_df.reset_index(drop=True))

        else:
            # --- No date column: forecast on index ---
            st.info(t('forecast_no_date'))
            series = df[fc_col].dropna().astype(float)
            if series.shape[0] < 3:
                st.warning(t('forecast_no_data'))
                return

            n = series.shape[0]
            deg = 1 if n < 6 else 2
            x = np.arange(n)
            coeffs = np.polyfit(x, series.values, deg)
            model = np.poly1d(coeffs)
            
            fitted = model(x)
            resid = series.values - fitted
            resid_std = np.nanstd(resid)
            ci = 1.96 * resid_std
            
            future_x = np.arange(n, n + int(fc_periods))
            preds = model(future_x)
            
            forecast_df = pd.DataFrame({
                'index': future_x,
                'forecast': preds,
                'lower_band': preds - ci,
                'upper_band': preds + ci
            })

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=x, y=series.values, mode='lines', name=t('actual')))
            fig.add_trace(go.Scatter(x=future_x, y=preds, mode='lines', name=t('forecast'), line=dict(dash='dash', color='red', width=3)))
            fig.add_trace(go.Scatter(
                x=list(future_x) + list(future_x[::-1]),
                y=list(preds + ci) + list(preds - ci)[::-1],
                fill='toself', fillcolor='rgba(255,0,0,0.15)',
                line=dict(color='rgba(255,255,255,0)'),
                hoverinfo="skip", showlegend=True, name=t('confidence')
            ))
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(forecast_df)

    except Exception as e:
        st.error(f"{t('forecast_fail')}: {e}")

# ================================================
# 5. EXPORTING HELPERS
# ================================================

def df_to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    """Converts a dictionary of DataFrames to an Excel file in bytes."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for name, df_sheet in sheets.items():
            if not isinstance(df_sheet, pd.DataFrame):
                continue
            safe_name = str(name)[:31]  # Excel sheet name limit
            df_sheet.to_excel(writer, sheet_name=safe_name, index=isinstance(df_sheet.index, pd.MultiIndex))
    out.seek(0)
    return out.getvalue()

def create_html_report(df: pd.DataFrame, insights: List[str]) -> bytes:
    """Generates a simple HTML report."""
    html = f'<html><head><meta charset="utf-8"><title>{t("title")}</title></head><body>'
    html += f'<h1>{t("title")}</h1>'
    html += f'<p>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>'
    html += f'<h2>Dataset</h2><p>{t("rows")}: {df.shape[0]} | {t("cols")}: {df.shape[1]}</p>'
    html += f'<h3>{t("insights")}</h3><ul>'
    for ins in insights:
        html += f'<li>{ins}</li>'
    html += '</ul>'
    html += f'<h3>{t("show_data")}</h3>'
    html += df.head(100).to_html(classes='table', border=1, justify='center')
    html += '</body></html>'
    return html.encode('utf-8')

def generate_pdf_report(df: pd.DataFrame, stats: pd.DataFrame, insights: List[str]) -> bytes:
    """Generates a professional PDF report with tables."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
    styles = getSampleStyleSheet()
    story = []

    # Title
    story.append(Paragraph(t('title'), styles['h1']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Spacer(1, 24))

    # Insights
    story.append(Paragraph(t('insights'), styles['h2']))
    for ins in insights:
        story.append(Paragraph(f"• {ins}", styles['Normal']))
    story.append(Spacer(1, 24))

    # Statistics
    if not stats.empty:
        story.append(Paragraph(t('stats_summary'), styles['h2']))
        stats_df_reset = stats.reset_index().rename(columns={'index': 'Metric'})
        stats_data = [stats_df_reset.columns.to_list()] + stats_df_reset.values.tolist()
        
        # Format numbers in data
        for i in range(1, len(stats_data)):
            for j in range(1, len(stats_data[i])):
                try:
                    stats_data[i][j] = f"{stats_data[i][j]:.2f}"
                except (TypeError, ValueError):
                    pass
        
        t_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        
        stats_table = Table(stats_data, colWidths=[1.5*inch] + [0.8*inch]*(len(stats_df_reset.columns)-1))
        stats_table.setStyle(t_style)
        story.append(stats_table)
        story.append(Spacer(1, 24))

    # Raw Data (Preview)
    story.append(Paragraph(t('show_data') + " (Top 50 rows)", styles['h2']))
    
    # Truncate data if too wide
    max_cols = 8
    df_preview = df.head(50)
    if df_preview.shape[1] > max_cols:
        df_preview = df_preview.iloc[:, :max_cols]
        story.append(Paragraph(f"(Showing first {max_cols} columns)", styles['Italic']))

    data = [df_preview.columns.to_list()] + df_preview.astype(str).values.tolist()
    
    t_style_data = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
    ])
    
    data_table = Table(data)
    data_table.setStyle(t_style_data)
    story.append(data_table)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ================================================
# 6. AUTOMATED INSIGHTS FUNCTION (WITH CACHING)
# ================================================

@st.cache_data
def get_automated_insights(df: pd.DataFrame) -> Tuple[List[str], Dict[str, str], Optional[str], Optional[str]]:
    """Generates a list of textual insights based on column names."""
    insights = []
    insights_dict = {}

    def safe_find(df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        for name in possible_names:
            for col in df.columns:
                if str(col).strip().lower() == str(name).strip().lower():
                    return col
        return None

    # Detect key columns
    revenue_col = safe_find(df, ["القيمة بعد الضريبة", "صافي المبيعات", "الإيرادات", "revenue", "total revenue", "sales"])
    discount_col = safe_find(df, ["الخصومات", "خصم", "discount", "total discount"])
    tax_col = safe_find(df, ["الضريبة", "ضريبة الصنف", "tax", "total tax"])
    qty_col = safe_find(df, ["الكمية", "كمية كرتون", "quantity", "total quantity"])
    branch_col = safe_find(df, ["الفرع", "branch"])
    salesman_col = safe_find(df, ["اسم المندوب", "مندوب", "salesman"])
    product_col = safe_find(df, ["اسم الصنف", "الصنف", "product", "category"])

    # Calculate totals
    if revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        total_revenue = df[revenue_col].sum()
        insights_dict["Total Revenue"] = f"{total_revenue:,.2f}"
        insights.append(f"💰 Total Revenue: {total_revenue:,.2f}")
    if discount_col and pd.api.types.is_numeric_dtype(df[discount_col]):
        total_discount = df[discount_col].sum()
        insights_dict["Total Discounts"] = f"{total_discount:,.2f}"
        insights.append(f"🎯 Total Discounts: {total_discount:,.2f}")
    if tax_col and pd.api.types.is_numeric_dtype(df[tax_col]):
        total_tax = df[tax_col].sum()
        insights_dict["Total Tax"] = f"{total_tax:,.2f}"
        insights.append(f"💸 Total Tax: {total_tax:,.2f}")
    if qty_col and pd.api.types.is_numeric_dtype(df[qty_col]):
        total_qty = df[qty_col].sum()
        insights_dict["Total Quantity"] = f"{total_qty:,.2f}"
        insights.append(f"📦 Total Quantity: {total_qty:,.2f}")

    # Find top categories
    if branch_col and revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        top_branch = df.groupby(branch_col)[revenue_col].sum().idxmax()
        insights_dict["Top Branch"] = str(top_branch)
        insights.append(f"🏢 Top Branch by Revenue: {top_branch}")
    if salesman_col and revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        top_salesman = df.groupby(salesman_col)[revenue_col].sum().idxmax()
        insights_dict["Top Salesman"] = str(top_salesman)
        insights.append(f"🧍‍♂️ Top Salesman: {top_salesman}")
    if product_col and revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        top_product = df.groupby(product_col)[revenue_col].sum().idxmax()
        insights_dict["Top Product"] = str(top_product)
        insights.append(f"🛒 Top Product: {top_product}")

    return insights, insights_dict, revenue_col, branch_col

# ================================================
# 7. MAIN STREAMLIT APP LAYOUT
# ================================================

def main():
    
    # --- Sidebar ---
    with st.sidebar:
        st.header(t('title'))
        lang_options = ['English', 'Arabic']
        lang_index = 1 if st.session_state.get('lang', 'en') == 'ar' else 0
        lang = st.selectbox(t('language'), options=lang_options, index=lang_index)
        st.session_state['lang'] = 'ar' if lang == 'Arabic' else 'en'
        
        dark = st.checkbox(t('theme'))
        if dark:
            st.markdown("""
            <style>
            .stApp { background-color: #0f1724; color: #e6edf3; }
            </style>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        uploaded_file = st.file_uploader(t('upload'), type=['xlsx', 'xls', 'csv', 'pdf', 'html', 'htm'])
        if uploaded_file:
            # Check if it's a new file before reloading
            if uploaded_file.name != st.session_state.get('file_name'):
                with st.spinner('Loading data...'):
                    load_data(uploaded_file)
        
        if st.button(t('load_sample')):
            with st.spinner('Loading sample data...'):
                load_sample_data()

    # --- Main Page ---
    st.title(t('title'))

    df = st.session_state.get('df')

    if df is None:
        st.info(f"ℹ️ {t('upload_prompt')}")
        
        # Footer
        st.markdown(
            """
            <hr style="margin-top:50px; margin-bottom:10px; border:1px solid #444;">
            <div style='text-align: center; color: #aaa; font-size: 14px;'>
                {t('footer_credit')} <b style='color:#00BFFF;'>Sameh Sobhy Attia</b>
            </div>
            """.replace('{t(\'footer_credit\')}', t('footer_credit')),
            unsafe_allow_html=True
        )
        return

    # --- Data Loaded - Show Tabs ---
    
    if st.checkbox(t('show_data')):
        st.dataframe(df)

    all_cols = df.columns.tolist()
    default_numeric = [c for c in all_cols if pd.api.types.is_numeric_dtype(df[c])]
    default_date = next((c for c in all_cols if 'date' in str(c).lower() or 'مبيعات' in str(c).lower()), None)
    date_col_index = all_cols.index(default_date) + 1 if default_date else 0
    
    # --- Tabbed Interface ---
    tab_kpi, tab_pivot, tab_charts, tab_forecast, tab_insights, tab_export = st.tabs([
        f"📊 {t('kpi_tab')}",
        f"📋 {t('pivot_tab')}",
        f"📈 {t('charts_tab')}",
        f"🔮 {t('forecast_tab')}",
        f"💡 {t('insights_tab')}",
        f"📄 {t('export_tab')}"
    ])

    # --- 1. KPI & Stats Tab ---
    with tab_kpi:
        st.subheader(t('config'))
        c1, c2 = st.columns(2)
        with c1:
            # This selection is used by other tabs (Forecast)
            date_col = st.selectbox(t('date_column'), options=[''] + all_cols, index=date_col_index, key='date_col_selector')
            date_col = date_col if date_col else None
        with c2:
            numeric_cols = st.multiselect(t('kpi_selection'), options=all_cols, default=default_numeric[:3])
        
        st.markdown("---")
        
        st.subheader(f"🔹 {t('total_everything')}")
        # Use cached function
        totals_dict_all, grand_all = grand_totals(df)
        kpi_cols_display = list(totals_dict_all.keys())[:5] # Show up to 5
        kpi_cols = st.columns(len(kpi_cols_display) if kpi_cols_display else 1)
        for i, k in enumerate(kpi_cols_display):
            kpi_cols[i].metric(k, f"{totals_dict_all[k]:,.2f}")
        st.metric(t('grand_total'), f"{grand_all:,.2f}")
        
        st.markdown("---")
        
        st.subheader(f"🔸 {t('selected_kpis')}")
        if numeric_cols:
            selected_df = df[numeric_cols].select_dtypes(include=[np.number])
            if not selected_df.empty:
                # This is a fast operation, no need to cache
                totals_dict = selected_df.sum(numeric_only=True).to_dict()
                grand_selected = selected_df.sum(numeric_only=True).sum()
                
                kpi_cols_sel = st.columns(len(totals_dict) if totals_dict else 1)
                for i, (col, val) in enumerate(totals_dict.items()):
                    kpi_cols_sel[i].metric(col, f"{val:,.2f}")
                st.metric(t('grand_total'), f"{grand_selected:,.2f}")
            else:
                st.info(t('no_kpis_selected'))
        else:
            st.info(t('no_kpis_selected'))

        st.markdown("---")
        st.subheader(t('stats_summary'))
        # Use cached function
        stat_df = stats_summary(df)
        if not stat_df.empty:
            st.dataframe(stat_df.style.format("{:,.2f}"))
        else:
            st.info(t('no_numeric_stats'))

    # --- 2. Pivot Table Tab ---
    with tab_pivot:
        st.subheader(t('pivot_config'))
        p1, p2 = st.columns(2)
        with p1:
            pivot_rows = st.multiselect(t('row_field'), options=all_cols, default=all_cols[0] if all_cols else [], key='pivot_rows')
            pivot_cols = st.multiselect(t('col_field'), options=all_cols, key='pivot_cols')
        with p2:
            pivot_value = st.selectbox(t('value_col'), options=[''] + all_cols, index=0, key='pivot_val')
            pivot_agg = st.selectbox(t('agg_type'), options=['sum', 'mean', 'median', 'count', 'min', 'max', 'std'], index=0, key='pivot_agg')
        
        if st.button(t('generate_pivot')):
            with st.spinner('Generating pivot table...'):
                pivot_value_arg = pivot_value if pivot_value else None
                if not pivot_value_arg:
                    pivot_agg = 'count'
                
                # Use cached function
                pvt = generate_pivot(df, rows=pivot_rows, cols=pivot_cols, values=pivot_value_arg, aggfunc=pivot_agg)
                
                if pvt is not None:
                    st.dataframe(pvt.style.format("{:,.2f}").background_gradient(cmap='viridis', axis=1))
                    
                    excel_bytes = df_to_excel_bytes({'pivot': pvt})
                    st.download_button(t('download_pivot'), data=excel_bytes, file_name='pivot_table.xlsx')
                else:
                    st.error("Could not generate pivot table. Check selections.")

    # --- 3. Charts Tab ---
    with tab_charts:
        st.subheader(t('charts'))
        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            chart_type = st.selectbox(t('chart_type'), options=['Line', 'Bar', 'Area', 'Scatter', 'Box', 'Pie', 'Heatmap'], key='chart_type')
        with ch2:
            x_axis = st.selectbox(t('x_axis'), options=[''] + all_cols, index=date_col_index, key='chart_x')
        with ch3:
            y_axes = st.multiselect(t('y_axis'), options=all_cols, default=default_numeric[:1], key='chart_y')

        if st.button(t('plot')):
            if not y_axes and chart_type not in ['Heatmap']:
                st.warning(t('plot_warn'))
            else:
                with st.spinner('Plotting...'):
                    try:
                        if chart_type in ['Line', 'Bar', 'Area', 'Scatter']:
                            x_arg = x_axis if x_axis else None
                            if x_arg:
                                df_melted = df.melt(id_vars=[x_arg], value_vars=y_axes, var_name='Metric', value_name='Value')
                            else:
                                df_melted = df[y_axes].melt(var_name='Metric', value_name='Value')
                                
                            if chart_type == 'Line':
                                fig = px.line(df_melted, x=x_arg, y='Value', color='Metric', title=f"{chart_type} Chart")
                            elif chart_type == 'Bar':
                                fig = px.bar(df_melted, x=x_arg, y='Value', color='Metric', title=f"{chart_type} Chart", barmode='group')
                            elif chart_type == 'Area':
                                fig = px.area(df_melted, x=x_arg, y='Value', color='Metric', title=f"{chart_type} Chart")
                            elif chart_type == 'Scatter':
                                fig = px.scatter(df_melted, x=x_arg, y='Value', color='Metric', title=f"{chart_type} Chart")
                            st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'Box':
                            fig = px.box(df[y_axes], y=y_axes)
                            st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'Pie':
                            names_col = x_axis if x_axis else (all_cols[0] if all_cols else None)
                            if names_col:
                                fig = px.pie(df, names=names_col, values=y_axes[0], title=f"Pie Chart: {y_axes[0]}")
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.warning("Please select an X-Axis for Pie chart labels.")
                        
                        elif chart_type == 'Heatmap':
                            num_df = df.select_dtypes(include=[np.number])
                            if num_df.shape[1] < 2:
                                st.warning(t('no_corr'))
                            else:
                                corr = num_df.corr()
                                fig = px.imshow(corr, text_auto=True, aspect="auto", title="Correlation Heatmap")
                                st.plotly_chart(fig, use_container_width=True)

                    except Exception as e:
                        st.error(f"Could not plot: {e}")

    # --- 4. Forecasting Tab ---
    with tab_forecast:
        st.subheader(t('forecasting'))
        fc1, fc2 = st.columns(2)
        with fc1:
            fc_col = st.selectbox(t('forecast_column'), options=[''] + default_numeric, index=0, key='fc_col')
        with fc2:
            fc_periods = st.number_input(t('forecast_periods'), min_value=1, max_value=365, value=12, key='fc_periods')
        
        if st.button(t('run_forecast')):
            with st.spinner('Running forecast...'):
                # Pass the globally selected date_col from the KPI tab
                run_forecast(df, date_col, fc_col, fc_periods)

    # --- 5. Data Insights Tab ---
    with tab_insights:
        st.subheader(t('insights'))
        with st.spinner('Generating insights...'):
            # Use cached function
            insights, insights_dict, rev_col, br_col = get_automated_insights(df)
            
            if insights_dict:
                c1, c2 = st.columns(2)
                with c1:
                    st.dataframe(pd.DataFrame(list(insights_dict.items()), columns=["Metric", "Value"]))
                with c2:
                    for ins in insights:
                        st.markdown(f"- {ins}")
                
                if rev_col and br_col and pd.api.types.is_numeric_dtype(df[rev_col]):
                    try:
                        st.markdown("---")
                        st.subheader(f"Revenue by {br_col}")
                        df_grouped = df.groupby(br_col, as_index=False)[rev_col].sum()
                        fig = px.bar(df_grouped, x=br_col, y=rev_col,
                                     title=f"Branch Performance", color=br_col, text_auto=".2s")
                        fig.update_layout(showlegend=False)
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception as e:
                        st.warning(f"Could not plot branch insights: {e}")
            else:
                st.info("No specific insights found for columns like 'Revenue', 'Branch', etc.")

        st.markdown("---")
        st.subheader(t('missing_values'))
        miss = df.isna().sum()
        miss = miss[miss > 0]
        if miss.empty:
            st.success("No missing values found.")
        else:
            st.dataframe(miss)

        st.markdown("---")
        st.subheader(t('correlations'))
        num_df = df.select_dtypes(include=[np.number])
        if num_df.shape[1] >= 2:
            st.dataframe(num_df.corr().style.background_gradient(cmap='vlag', vmin=-1, vmax=1).format("{:,.2f}"))
        else:
            st.info(t('no_corr'))

    # --- 6. Export Tab ---
    with tab_export:
        st.subheader(t('export_tab'))
        # Get cached insights and stats
        insights, _, _, _ = get_automated_insights(df)
        stat_df = stats_summary(df)

        # Excel Download
        excel_data = df_to_excel_bytes({
            'Raw_Data': df,
            'Statistics': stat_df.reset_index()
        })
        st.download_button(
            label=f"📥 {t('download_excel')}",
            data=excel_data,
            file_name=f"Sales_Summary_{st.session_state.get('file_name', 'report')}.xlsx",
            mime="application/vnd.ms-excel"
        )
        
        # HTML Download
        html_data = create_html_report(df, insights)
        st.download_button(
            label=f"📥 {t('download_html')}",
            data=html_data,
            file_name=f"Sales_Report_{st.session_state.get('file_name', 'report')}.html",
            mime="text/html"
        )
        
        # PDF Download
        try:
            with st.spinner('Generating PDF Report...'):
                pdf_data = generate_pdf_report(df, stat_df, insights)
            st.download_button(
                label=f"📥 {t('download_pdf')}",
                data=pdf_data,
                file_name=f"Sales_Report_{st.session_state.get('file_name', 'report')}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Could not generate PDF. Error: {e}")

    # --- Footer ---
    st.markdown(
        """
        <hr style="margin-top:50px; margin-bottom:10px; border:1px solid #444;">
        <div style='text-align: center; color: #aaa; font-size: 14px;'>
            {t('footer_credit')} <b style='color:#00BFFF;'>Sameh Sobhy Attia</b> (Pro Version by Gemini)
        </div>
        """.replace('{t(\'footer_credit\')}', t('footer_credit')),
        unsafe_allow_html=True
    )

# ================================================
# RUN THE APP
# ================================================
if __name__ == "__main__":
    main()

