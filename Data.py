# Sales_Insights_Pro.py
#
# A professional, multi-lingual, multi-file-type Sales Dashboard and Forecasting tool.
# Version 4.0: Added simple, session-based username/password authentication.
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
# - Simple Authentication (Username: Tiba, Password: Hamza321*).
# - Caching for high-performance data processing.
# - Interactive Dashboard: Click/select rows to dynamically update charts.
# - Supports Excel, CSV, PDF, and HTML (table extraction) file uploads.
# - Fully bilingual (English/Arabic) UI.
# - Robust session state management (data persists across interactions).
# - Clean, tabbed interface.
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
if 'authenticated' not in st.session_state: # NEW: Authentication flag
    st.session_state['authenticated'] = False # NEW
if 'login_user' not in st.session_state: # FIX: Initialize login_user to prevent StreamlitAPIException
    st.session_state['login_user'] = None # FIX

# Hardcoded credentials for demonstration
CREDENTIALS = {
    "username": "Tiba",
    "password": "Hamza321*"
}

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
        'dashboard_tab': 'Interactive Dashboard',
        'pivot_tab': 'Pivot Table',
        'charts_tab': 'Manual Charts',
        'forecast_tab': 'Forecasting',
        'insights_tab': 'Data Insights',
        'export_tab': 'Export Report',
        'selected_kpis': 'Totals for Selected KPIs',
        'no_kpis_selected': 'No KPI columns selected.',
        'no_numeric_stats': 'No numeric columns for statistics.',
        'plot_warn': 'Please select at least one Y-Axis column.',
        'forecast_warn': 'Please select a numeric column to forecast.',
        'forecast_no_date': 'No date column selected. Forecasting on data index.',
        'forecast_no_data': 'Not enough data to forecast (need at least 2 data points).',
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
        'dashboard_info': 'Select rows from the table below to dynamically generate charts based on your selection.',
        'plot_selection_title': 'Plot for Selected Data',
        'plot_all_title': 'Plot for All Data (No Rows Selected)',
        'stat_metric': 'Metric',
        'stat_value': 'Value',
        'stat_count': 'Count',
        'stat_mean': 'Average',
        'stat_median': 'Median',
        'stat_max': 'Max',
        'stat_min': 'Min',
        'stat_std': 'Std. Dev.',
        'insight_total_revenue': 'Total Revenue',
        'insight_total_discounts': 'Total Discounts',
        'insight_total_tax': 'Total Tax',
        'insight_total_qty': 'Total Quantity',
        'insight_top_branch': 'Top Branch',
        'insight_top_salesman': 'Top Salesman',
        'insight_top_product': 'Top Product',
        # NEW AUTH TRANSLATIONS
        'login': 'Login',
        'username': 'Username',
        'password': 'Password',
        'logout': 'Logout',
        'login_title': 'Please Log In to Access the Dashboard',
        'login_error': 'Incorrect username or password. Please try again.',
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
        'dashboard_tab': 'لوحة تحكم تفاعلية',
        'pivot_tab': 'الجدول المحوري',
        'charts_tab': 'مخططات يدوية',
        'forecast_tab': 'التنبؤ',
        'insights_tab': 'تحليلات البيانات',
        'export_tab': 'تصدير التقرير',
        'selected_kpis': 'إجماليات المؤشرات المحددة',
        'no_kpis_selected': 'لم يتم تحديد أعمدة مؤشرات.',
        'no_numeric_stats': 'لا توجد أعمدة رقمية للإحصاءات.',
        'plot_warn': 'يرجى اختيار عمود واحد على الأقل للمحور الصادي.',
        'forecast_warn': 'يرجى اختيار عمود رقمي للتنبؤ.',
        'forecast_no_date': 'لم يتم تحديد عمود تاريخ. سيتم التنبؤ بناءً على تسلسل البيانات.',
        'forecast_no_data': 'لا توجد بيانات كافية للتنبؤ (تحتاج نقطتي بيانات على الأقل).',
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
        'dashboard_info': 'اختر صفوفاً من الجدول أدناه لإنشاء مخططات ديناميكياً بناءً على اختيارك.',
        'plot_selection_title': 'مخطط للبيانات المحددة',
        'plot_all_title': 'مخطط لكل البيانات (لم يتم تحديد صفوف)',
        'stat_metric': 'المقياس',
        'stat_value': 'القيمة',
        'stat_count': 'العدد',
        'stat_mean': 'المتوسط',
        'stat_median': 'الوسيط',
        'stat_max': 'الأعلى',
        'stat_min': 'الأدنى',
        'stat_std': 'الانحراف المعياري',
        'insight_total_revenue': 'إجمالي الإيرادات',
        'insight_total_discounts': 'إجمالي الخصومات',
        'insight_total_tax': 'إجمالي الضريبة',
        'insight_total_qty': 'إجمالي الكمية',
        'insight_top_branch': 'أفضل فرع',
        'insight_top_salesman': 'أفضل بائع',
        'insight_top_product': 'أفضل منتج',
        # NEW AUTH TRANSLATIONS
        'login': 'تسجيل الدخول',
        'username': 'اسم المستخدم',
        'password': 'كلمة المرور',
        'logout': 'تسجيل الخروج',
        'login_title': 'الرجاء تسجيل الدخول للوصول إلى لوحة التحكم',
        'login_error': 'اسم المستخدم أو كلمة المرور غير صحيح. الرجاء المحاولة مرة أخرى.',
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
            
            # UPDATED: Allow forecast for 2 points (for a straight line)
            if tmp_series.shape[0] < 2:
                st.warning(t('forecast_no_data'))
                return

            n = tmp_series.shape[0]
            deg = 1 # Always use degree 1 (straight line) if n < 6
            if n >= 6:
                deg = 2 # Use degree 2 (curve) if 6 or more points
                
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
            # UPDATED: Allow forecast for 2 points (for a straight line)
            if series.shape[0] < 2:
                st.warning(t('forecast_no_data'))
                return

            n = series.shape[0]
            deg = 1 # Always use degree 1 (straight line) if n < 6
            if n >= 6:
                deg = 2 # Use degree 2 (curve) if 6 or more points
                
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
        # Use translated key for the index column
        stats_df_reset = stats.reset_index().rename(columns={'index': t('stat_metric')})
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
def get_automated_insights(df: pd.DataFrame) -> Tuple[List[Tuple[str, str, str]], Dict[str, str], Optional[str], Optional[str]]:
    """Generates a list of textual insights based on column names."""
    # Insights is now a list of tuples (emoji, key, value)
    insights: List[Tuple[str, str, str]] = []
    insights_dict = {}

    def safe_find(df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        for name in possible_names:
            for col in df.columns:
                if str(col).strip().lower() == str(name).strip().lower():
                    return col
        return None

    # Detect key columns
    revenue_col = safe_find(df, ["القيمة بعد الضريبة", "صافي المبيعات", "الإيرادات", "revenue", "total revenue", "sales", "قيمة بعد الضريبة و الخصم"])
    discount_col = safe_find(df, ["الخصومات", "خصم", "discount", "total discount"])
    tax_col = safe_find(df, ["الضريبة", "ضريبة الصنف", "tax", "total tax"])
    qty_col = safe_find(df, ["الكمية", "كمية كرتون", "quantity", "total quantity"])
    branch_col = safe_find(df, ["الفرع", "branch"])
    salesman_col = safe_find(df, ["اسم المندوب", "مندوب", "salesman", "seller", "بائع"])
    product_col = safe_find(df, ["اسم الصنف", "الصنف", "product", "category"])

    # Calculate totals
    if revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        total_revenue = df[revenue_col].sum()
        insights_dict['insight_total_revenue'] = f"{total_revenue:,.2f}"
        insights.append(('💰', 'insight_total_revenue', f"{total_revenue:,.2f}"))
    if discount_col and pd.api.types.is_numeric_dtype(df[discount_col]):
        total_discount = df[discount_col].sum()
        insights_dict['insight_total_discounts'] = f"{total_discount:,.2f}"
        insights.append(('🎯', 'insight_total_discounts', f"{total_discount:,.2f}"))
    if tax_col and pd.api.types.is_numeric_dtype(df[tax_col]):
        total_tax = df[tax_col].sum()
        insights_dict['insight_total_tax'] = f"{total_tax:,.2f}"
        insights.append(('💸', 'insight_total_tax', f"{total_tax:,.2f}"))
    if qty_col and pd.api.types.is_numeric_dtype(df[qty_col]):
        total_qty = df[qty_col].sum()
        insights_dict['insight_total_qty'] = f"{total_qty:,.2f}"
        insights.append(('📦', 'insight_total_qty', f"{total_qty:,.2f}"))

    # Find top categories
    if branch_col and revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        top_branch = df.groupby(branch_col)[revenue_col].sum().idxmax()
        insights_dict['insight_top_branch'] = str(top_branch)
        insights.append(('🏢', 'insight_top_branch', str(top_branch)))
    if salesman_col and revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        top_salesman = df.groupby(salesman_col)[revenue_col].sum().idxmax()
        insights_dict['insight_top_salesman'] = str(top_salesman)
        insights.append(('🧍‍♂️', 'insight_top_salesman', str(top_salesman)))
    if product_col and revenue_col and pd.api.types.is_numeric_dtype(df[revenue_col]):
        top_product = df.groupby(product_col)[revenue_col].sum().idxmax()
        insights_dict['insight_top_product'] = str(top_product)
        insights.append(('🛒', 'insight_top_product', str(top_product)))

    return insights, insights_dict, revenue_col, branch_col

# ================================================
# 7. DYNAMIC PLOTTING FUNCTION (FOR DASHBOARD)
# ================================================

def plot_dynamic_chart(data: pd.DataFrame, chart_type: str, x_axis: Optional[str], y_axes: List[str]):
    """Helper function to generate plots for the interactive dashboard."""
    if not y_axes and chart_type not in ['Heatmap']:
        st.warning(t('plot_warn'))
        return
    
    try:
        if chart_type in ['Line', 'Bar', 'Area', 'Scatter']:
            x_arg = x_axis if x_axis else None
            if x_arg:
                df_melted = data.melt(id_vars=[x_arg], value_vars=y_axes, var_name='Metric', value_name='Value')
            else:
                df_melted = data[y_axes].melt(var_name='Metric', value_name='Value')
                
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
            fig = px.box(data[y_axes], y=y_axes)
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == 'Pie':
            names_col = x_axis if x_axis else (data.columns[0] if not data.empty else None)
            if names_col and y_axes:
                fig = px.pie(data, names=names_col, values=y_axes[0], title=f"Pie Chart: {y_axes[0]}")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Please select an X-Axis (for labels) and at least one Y-Axis (for values).")
        
        elif chart_type == 'Heatmap':
            num_df = data.select_dtypes(include=[np.number])
            if num_df.shape[1] < 2:
                st.warning(t('no_corr'))
            else:
                corr = num_df.corr()
                fig = px.imshow(corr, text_auto=True, aspect="auto", title="Correlation Heatmap")
                st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"Could not plot: {e}")

# ================================================
# 8. AUTHENTICATION & APP CONTENT
# ================================================

def login_form():
    """Displays the login form."""
    st.image("https://placehold.co/600x100/1e293b/FFFFFF?text=SALES+INSIGHTS", width=400)
    st.title(t('login_title'))

    # Center the login form
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input(t('username'), key="login_user_input")
            password = st.text_input(t('password'), type="password", key="login_pass_input")
            submitted = st.form_submit_button(t('login'), type="primary")

            if submitted:
                if username == CREDENTIALS['username'] and password == CREDENTIALS['password']:
                    st.session_state['authenticated'] = True
                    st.session_state['login_user'] = username # Store username
                    st.rerun()
                else:
                    st.error(t('login_error'))
    
    # Display the footer even when logged out
    st.markdown(
        """
        <hr style="margin-top:50px; margin-bottom:10px; border:1px solid #444;">
        <div style='text-align: center; color: #aaa; font-size: 14px;'>
            {t_footer_credit} <b style='color:#00BFFF;'>Sameh Sobhy Attia</b> (Pro Version by Gemini)
        </div>
        """.replace('{t_footer_credit}', t('footer_credit')),
        unsafe_allow_html=True
    )


def app_content():
    """Contains all the original application logic (tabs, analysis, etc.)."""
    
    # Header with Logout and Language/Theme toggle
    c_header_1, c_header_2, c_header_3, c_header_4 = st.columns([1, 1, 1, 1])
    
    with c_header_1:
        st.title(t('title'))
    
    with c_header_2:
        pass # Placeholder for layout
        
    with c_header_3:
        # Language Selector
        if st.session_state['lang'] == 'en':
            selected_lang = st.selectbox(t('language'), options=['English', 'العربية'], index=0, key='lang_select')
        else:
            selected_lang = st.selectbox(t('language'), options=['English', 'العربية'], index=1, key='lang_select')
            
        if selected_lang == 'العربية' and st.session_state['lang'] != 'ar':
            st.session_state['lang'] = 'ar'
            st.rerun()
        elif selected_lang == 'English' and st.session_state['lang'] != 'en':
            st.session_state['lang'] = 'en'
            st.rerun()
            
    with c_header_4:
        # Logout Button
        if st.button(t('logout')):
            st.session_state['authenticated'] = False
            st.session_state['login_user'] = None
            st.success("Logged out successfully!")
            st.rerun()
            
    # Check current language for RTL layout (Arabic)
    if st.session_state['lang'] == 'ar':
        st.markdown("<style>body { direction: rtl; text-align: right; }</style>", unsafe_allow_html=True)
        # Apply RTL to selectboxes and inputs within Streamlit components (best effort)
        st.markdown("""
        <style>
        .stSelectbox, .stTextInput, .stMultiSelect {
            direction: rtl;
        }
        </style>
        """, unsafe_allow_html=True)

    df = st.session_state.get('df')

    # --- Data Loading Controls ---
    with st.sidebar:
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
        
    if df is None:
        st.info(f"ℹ️ {t('upload_prompt')}")
        return

    # --- Data Loaded - Show Tabs ---
    
    if st.checkbox(t('show_data')):
        # Calculate height: (rows + 1 header) * 35px/row + 3px extra
        table_height = (len(df) + 1) * 35 + 3
        # Set a max height to avoid crashing the browser on huge datasets
        if table_height > 1000:
            table_height = 1000
        st.dataframe(df, use_container_width=True, height=table_height)

    all_cols = df.columns.tolist()
    default_numeric = [c for c in all_cols if pd.api.types.is_numeric_dtype(df[c])]
    default_date = next((c for c in all_cols if 'date' in str(c).lower() or 'تاريخ' in str(c).lower()), None)
    date_col_index = all_cols.index(default_date) + 1 if default_date else 0
    
    # --- Tabbed Interface ---
    tab_kpi, tab_dashboard, tab_pivot, tab_charts, tab_forecast, tab_insights, tab_export = st.tabs([
        f"📊 {t('kpi_tab')}",
        f"✨ {t('dashboard_tab')}",
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
            # UPDATED: Rename columns using translations
            stat_df = stat_df.rename(columns={
                'count': t('stat_count'),
                'mean': t('stat_mean'), # This becomes 'Average'
                'median': t('stat_median'),
                'max': t('stat_max'),
                'min': t('stat_min'),
                'std': t('stat_std')
            })
            st.dataframe(stat_df.style.format("{:,.2f}"))
        else:
            st.info(t('no_numeric_stats'))

    # --- 2. Interactive Dashboard Tab ---
    with tab_dashboard:
        st.subheader(t('dashboard_tab'))
        st.info(t('dashboard_info'))

        # --- Dashboard Controls ---
        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            dash_chart_type = st.selectbox(t('chart_type'), options=['Line', 'Bar', 'Area', 'Scatter', 'Box', 'Pie'], key='dash_chart_type')
        with ch2:
            dash_x_axis = st.selectbox(t('x_axis'), options=[''] + all_cols, index=date_col_index, key='dash_x')
        with ch3:
            dash_y_axes = st.multiselect(t('y_axis'), options=all_cols, default=default_numeric[:1], key='dash_y')

        # --- Interactive Dataframe ---
        st.dataframe(df, on_select="rerun", selection_mode="multi-row", key="dashboard_selector", use_container_width=True, height=300)

        # --- Check selection and plot ---
        selection_state = st.session_state.get("dashboard_selector", {})
        selected_rows_indices = selection_state.get("selection", {}).get("rows", [])

        if selected_rows_indices:
            selected_df = df.iloc[selected_rows_indices]
            st.subheader(f"{t('plot_selection_title')} ({len(selected_rows_indices)} {t('rows')})")
            plot_dynamic_chart(selected_df, dash_chart_type, dash_x_axis, dash_y_axes)
        else:
            st.subheader(t('plot_all_title'))
            plot_dynamic_chart(df, dash_chart_type, dash_x_axis, dash_y_axes)


    # --- 3. Pivot Table Tab ---
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
                pivot_value_arg = pivot_value if pivot_value else (all_cols[0] if all_cols else None)
                pivot_result = generate_pivot(df, pivot_rows, pivot_cols, pivot_value_arg, pivot_agg)
                if pivot_result is not None:
                    st.dataframe(pivot_result)
                    
                    # Download button for pivot table
                    excel_bytes = df_to_excel_bytes({'PivotTable': pivot_result})
                    st.download_button(
                        label=t('download_pivot'),
                        data=excel_bytes,
                        file_name=f"Pivot_Table_{st.session_state['file_name']}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    # --- 4. Manual Charts Tab ---
    with tab_charts:
        st.subheader(t('charts'))
        
        c_m_1, c_m_2, c_m_3 = st.columns(3)
        with c_m_1:
            manual_chart_type = st.selectbox(t('chart_type'), options=['Line', 'Bar', 'Area', 'Scatter', 'Box', 'Pie', 'Heatmap'], key='manual_chart_type')
        with c_m_2:
            manual_x_axis = st.selectbox(t('x_axis'), options=[''] + all_cols, index=date_col_index, key='manual_x')
        with c_m_3:
            manual_y_axes = st.multiselect(t('y_axis'), options=all_cols, default=default_numeric[:1], key='manual_y')

        if st.button(t('plot')):
            plot_dynamic_chart(df, manual_chart_type, manual_x_axis, manual_y_axes)

    # --- 5. Forecasting Tab ---
    with tab_forecast:
        st.subheader(t('forecasting'))
        
        c_f_1, c_f_2, c_f_3 = st.columns(3)
        with c_f_1:
            # Use the KPI selector for the date column
            # Note: date_col is defined in the KPI tab scope, so use the key
            fc_date_col = st.session_state.get('date_col_selector', None)
            st.markdown(f"**{t('date_column')}:** `{fc_date_col if fc_date_col else 'None'}`")
            
        with c_f_2:
            fc_col = st.selectbox(t('forecast_column'), options=[''] + default_numeric, key='fc_col')
        with c_f_3:
            fc_periods = st.number_input(t('forecast_periods'), min_value=1, max_value=36, value=6, step=1, key='fc_periods')

        if st.button(t('run_forecast'), type="primary"):
            with st.spinner('Calculating forecast...'):
                if fc_col:
                    run_forecast(df, fc_date_col, fc_col, fc_periods)
                else:
                    st.warning(t('forecast_warn'))

    # --- 6. Data Insights Tab ---
    with tab_insights:
        st.subheader(t('insights'))
        
        insights_list, insights_dict, revenue_col, branch_col = get_automated_insights(df)
        
        # Display insights as KPIs
        if insights_list:
            cols = st.columns(len(insights_list))
            for i, (emoji, key, value) in enumerate(insights_list):
                cols[i].metric(f"{emoji} {t(key)}", value)

        st.markdown("---")
        
        # Missing values
        st.subheader(t('missing_values'))
        missing = df.isnull().sum()
        missing = missing[missing > 0].sort_values(ascending=False)
        if not missing.empty:
            missing_df = pd.DataFrame({t('stat_metric'): missing.index, t('stat_count'): missing.values, '% Missing': (missing.values / len(df)) * 100})
            fig = px.bar(missing_df, x=t('stat_metric'), y=t('stat_count'), 
                         title=t('missing_values'), color='% Missing')
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(missing_df, hide_index=True)
        else:
            st.info("No missing values found in the dataset.")

        st.markdown("---")
        
        # Correlation Matrix
        st.subheader(t('correlations'))
        num_df = df.select_dtypes(include=[np.number])
        if num_df.shape[1] >= 2:
            corr = num_df.corr()
            fig = px.imshow(corr, text_auto=".2f", aspect="auto", title=t('correlations'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(t('no_corr'))


    # --- 7. Export Report Tab ---
    with tab_export:
        st.subheader(t('export_tab'))
        
        stats = stats_summary(df)
        insights_list_text = [f"{t(key)}: {value}" for _, key, value in get_automated_insights(df)[0]]
        
        st.markdown("### Export Options")
        
        # Export Excel (Summary)
        excel_sheets = {
            'Raw Data Preview': df.head(100),
            'Statistics Summary': stats if not stats.empty else pd.DataFrame(),
            'Totals': pd.DataFrame([totals_dict_all])
        }
        excel_data = df_to_excel_bytes(excel_sheets)
        st.download_button(
            label=t('download_excel'),
            data=excel_data,
            file_name=f"Sales_Summary_{st.session_state['file_name']}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Export HTML
        html_data = create_html_report(df, insights_list_text)
        st.download_button(
            label=t('download_html'),
            data=html_data,
            file_name=f"Sales_Report_{st.session_state['file_name']}_{datetime.now().strftime('%Y%m%d')}.html",
            mime="text/html"
        )
        
        # Export PDF
        pdf_data = generate_pdf_report(df, stats, insights_list_text)
        st.download_button(
            label=t('download_pdf'),
            data=pdf_data,
            file_name=f"Sales_Report_{st.session_state['file_name']}_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )

# ================================================
# 9. MAIN APP EXECUTION FLOW
# ================================================

# Main execution flow based on authentication state
if not st.session_state['authenticated']:
    login_form()
else:
    app_content()
    
# Footer (always visible in app content)
if st.session_state['authenticated']:
    st.markdown(
        """
        <hr style="margin-top:50px; margin-bottom:10px; border:1px solid #444;">
        <div style='text-align: center; color: #aaa; font-size: 14px;'>
            {t_footer_credit} <b style='color:#00BFFF;'>Sameh Sobhy Attia</b> (Pro Version by Gemini)
        </div>
        """.replace('{t_footer_credit}', t('footer_credit')),
        unsafe_allow_html=True
    )
