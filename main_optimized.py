import streamlit as st 
import os
import pandas as pd
from datetime import datetime
import plotly.express as px
import pdfkit
from jinja2 import Template
import uuid
import tempfile
from datetime import timedelta
import seaborn as sns
import matplotlib.pyplot as plt
import plotly.io as pio

# ==============================================================================
# ĐẢM BẢO FILE 'a04ecaf1_1dae_4c90_8081_086cd7c7b725.py' NẰNG CÙNG THƯ MỤC
# HOẶC THAY THẾ TÊN FILE NẾU BẠN ĐÃ ĐỔI TÊN NÓ.
# ==============================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report,
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

# Gọi hàm setup_paths ngay từ đầu để path_dict có sẵn
path_dict = setup_paths()

# ==============================================================================
# KHỞI TẠO CÁC BIẾN TRẠNG THÁI PHIÊN (SESSION STATE VARIABLES)
# ==============================================================================
if 'comparison_mode_select_tab_main' not in st.session_state:
    st.session_state.comparison_mode_select_tab_main = "Compare Projects in a Month"  
    
if 'comparison_mode' not in st.session_state:
    st.session_state.comparison_mode = "So Sánh Dự Án Trong Một Tháng" 

if 'comparison_selected_years' not in st.session_state:
    st.session_state.comparison_selected_years = [datetime.now().year] 

if 'comparison_selected_months' not in st.session_state:
    st.session_state.comparison_selected_months = [] 

if 'comparison_selected_projects' not in st.session_state:
    st.session_state.comparison_selected_projects = [] 

if 'comparison_selected_months_over_time' not in st.session_state:
    st.session_state.comparison_selected_months_over_time = [] 

if 'selected_years' not in st.session_state: 
    st.session_state.selected_years = [datetime.now().year]

if 'selected_months' not in st.session_state: 
    st.session_state.selected_months = []

if 'selected_language' not in st.session_state:
    st.session_state.selected_language = "English"

# ---------------------------
# PHẦN XÁC THỰC TRUY CẬP
# ---------------------------
@st.cache_data
def load_invited_emails():
    try:
        df = pd.read_csv(csv_file_path, header=None, encoding='utf-8')
        emails = df.iloc[:, 0].astype(str).str.strip().str.lower().tolist()
        return emails
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file invited_emails.csv tại {csv_file_path}. Vui lòng kiểm tra đường dẫn.")
        return []
    except Exception as e:
        st.error(f"Lỗi khi tải file invited_emails.csv: {e}")
        return []

INVITED_EMAILS = load_invited_emails()

def log_user_access(email):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {"Time": timestamp, "Email": email}
    if "access_log" not in st.session_state:
        st.session_state.access_log = []
    st.session_state.access_log.append(log_entry)

if "user_email" not in st.session_state:
    st.set_page_config(page_title="Triac Time Report", layout="wide")
    st.title("🔐 Access authentication")
    email_input = st.text_input("📧 Enter the invited email to access:")

    if email_input:
        email = email_input.strip().lower()
        if email in INVITED_EMAILS:
            st.session_state.user_email = email
            log_user_access(email)
            st.success("✅ Valid email! Entering application...")
            st.rerun()
        else:
            st.error("❌ Email is not on the invitation list.")
    st.stop() 

# ---------------------------
# PHẦN GIAO DIỆN CHÍNH CỦA ỨNG DỤNG
# ---------------------------
if 'lang' not in st.session_state:
    st.session_state.lang = 'en' 

st.set_page_config(page_title="Triac Time Report", layout="wide")

st.markdown("""
    <style>
        .report-title {font-size: 30px; color: #003366; font-weight: bold;}
        .report-subtitle {font-size: 14px; color: gray;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

TEXTS = {
    'en': {
        'app_title': "📊 Time Report Generator",
        'lang_select': "Select language:",
        'language_vi': "Tiếng Việt",
        'language_en': "English",
        'template_not_found': "❌ Template file not found: {}. Please ensure the file is in the same directory as the application.",
        'failed_to_load_raw_data': "⚠️ Failed to load raw data. Please check the 'Raw Data' sheet in the template file and data format.",
        'loading_data': "🔄 Loading data and configurations...",
        'tab_standard_report': "Standard Report",
        'tab_comparison_report': "Comparison Report",
        'tab_data_preview': "Data Preview",
        'standard_report_header': "Standard Time Report Configuration",
        'select_analysis_mode': "Select analysis mode:",
        'select_year': "Select year:",
        'select_months': "Select month(s):",
        'standard_project_selection_header': "Project Selection for Standard Report",
        'standard_project_selection_text': "Select projects to include (only 'yes' projects from template config will be included by default):",
        'generate_standard_report_btn': "🚀 Generate Standard Report",
        'no_year_selected_error': "Please select a valid year to generate the report.",
        'no_project_selected_warning_standard': "Please select at least one project to generate the standard report.",
        'no_data_after_filter_standard': "⚠️ No data after filtering for the standard report. Please check your selections.",
        'latest_update_date': "Latest data update",
        'generating_excel_report': "Generating Excel report...",
        'excel_report_generated': "✅ Excel Report generated: {}",
        'download_excel_report': "📥 Download Excel Report",
        'generating_pdf_report': "Generating PDF report...",
        'pdf_report_generated': "✅ PDF Report generated: {}",
        'download_pdf_report': "📥 Download PDF Report",
        'failed_to_generate_excel': "❌ Failed to generate Excel report.",
        'failed_to_generate_pdf': "❌ Failed to generate PDF report.",
        'comparison_report_header': "Comparison Report Configuration",
        'select_comparison_mode': "Select comparison mode:",
        'compare_projects_month': "Compare Projects in a Month",
        'compare_projects_year': "Compare Projects in a Year",
        'compare_projects_over_time': "Compare Projects Over Time (Months/Years)",
        'filter_data_for_comparison': "Filter Data for Comparison",
        'select_years': "Select Year(s):",
        'select_months_comp': "Select Month(s):",
        'select_projects_comp': "Select Project(s):",
        'generate_comparison_report_btn': "🚀 Generate Comparison Report",
        'no_data_after_filter_comparison': "⚠️ {}",
        'data_filtered_success': "✅ Data filtered successfully for comparison.",
        'comparison_data_preview': "Comparison Data Preview",
        'generating_comparison_excel': "Generating Comparison Excel Report...",
        'comparison_excel_generated': "✅ Comparison Excel Report generated: {}",
        'download_comparison_excel': "📥 Download Comparison Excel",
        'generating_comparison_pdf': "Generating Comparison PDF Report...",
        'comparison_pdf_generated': "✅ PDF Report generated: {}",
        'download_comparison_pdf': "📥 Download Comparison PDF",
        'failed_to_generate_comparison_excel': "❌ Failed to generate Comparison Excel report.",
        'failed_to_generate_comparison_pdf': "❌ Failed to generate Comparison PDF report.",
        'raw_data_preview_header': "Raw Input Data (First 100 rows)",
        'no_raw_data': "No raw data loaded.",
        'no_year_in_data': "No years in data to select.",
        'user_guide': "User Guide",
        'export_options': "Export Options",
        'export_excel_option': "Export as Excel (.xlsx)",
        'export_pdf_option': "Export as PDF (.pdf)",
        'report_button': "Generate report",
        'no_data': "No data after filtering",
        'report_done': "Report created successfully",
        'download_excel': "Download Excel",
        'download_pdf': "Download PDF",
        'warning_select_export_format': "Please select at least one report export format (Excel or PDF).",
        'error_generating_report': "An error occurred while generating the report. Please try again.",
        'select_at_least_two_projects_warning': "Please select at least two projects for comparison.",
        'select_years_for_over_time_months': "Select the year(s) for comparison:",
        'select_months_for_single_year': "Select month(s) within the chosen year:",
        'comparison_over_years_note': "Note: You have selected multiple years. The report will compare the project's data across the selected years. Month selection will be ignored.",
        'comparison_over_months_note': "Note: The report will compare the project's data across the selected months in year {}.",
        'no_comparison_criteria_selected': "Please select at least one year or month for comparison.",
        'no_month_selected_for_single_year': "Please select at least one month when comparing a single project within a specific year.",
        'tab_help': "Help",
        'preview_charts_title': "📊 Preview Charts",
        "help_instruction_simple": "If you have any questions or need support, please email to Admin **ky@triaccomposites.com**. We will respond as soon as possible. Thank you!",
        'select_all_projects_checkbox': "Select all projects"
    },
    'vi': {
        'app_title': "📊 Công cụ tạo báo cáo thời gian",
        'lang_select': "Chọn ngôn ngữ:",
        'language_vi': "Tiếng Việt",
        'language_en': "English",
        'template_not_found': "❌ Không tìm thấy file template: {}. Vui lòng đảm bảo file nằm cùng thư mục với ứng dụng.",
        'failed_to_load_raw_data': "⚠️ Không thể tải dữ liệu thô. Vui lòng kiểm tra sheet 'Raw Data' trong file template và định dạng dữ liệu.",
        'loading_data': "🔄 Đang tải dữ liệu và cấu hình...",
        'tab_standard_report': "Báo cáo tiêu chuẩn",
        'tab_comparison_report': "Báo cáo so sánh",
        'tab_data_preview': "Xem trước dữ liệu",
        'standard_report_header': "Cấu hình báo cáo thời gian tiêu chuẩn",
        'select_analysis_mode': "Chọn chế độ phân tích:",
        'select_year': "Chọn năm:",
        'select_months': "Chọn tháng(các tháng):",
        'standard_project_selection_header': "Lựa chọn dự án cho báo cáo tiêu chuẩn",
        'standard_project_selection_text': "Chọn dự án để bao gồm (mặc định chỉ bao gồm các dự án 'yes' từ cấu hình template):",
        'generate_standard_report_btn': "🚀 Tạo báo cáo tiêu chuẩn",
        'no_year_selected_error': "Vui lòng chọn một năm hợp lệ để tạo báo cáo.",
        'no_project_selected_warning_standard': "Vui lòng chọn ít nhất một dự án để tạo báo cáo tiêu chuẩn.",
        'no_data_after_filter_standard': "⚠️ Không có dữ liệu sau khi lọc cho báo cáo tiêu chuẩn. Vui lòng kiểm tra các lựa chọn của bạn.",
        'generating_excel_report': "Đang tạo báo cáo Excel...",
        'excel_report_generated': "✅ Báo cáo Excel đã được tạo: {}",
        'download_excel_report': "📥 Tải báo cáo Excel",
        'generating_pdf_report': "Đang tạo báo cáo PDF...",
        'pdf_report_generated': "✅ Báo cáo PDF đã được tạo: {}",
        'download_pdf_report': "📥 Tải báo cáo PDF",
        'failed_to_generate_excel': "❌ Đã xảy ra lỗi khi tạo báo cáo Excel.",
        'failed_to_generate_pdf': "❌ Đã xảy ra lỗi khi tạo báo cáo PDF.",
        'comparison_report_header': "Cấu hình báo cáo so sánh",
        'select_comparison_mode': "Chọn chế độ so sánh:",
        'compare_projects_month': "So Sánh Dự Án Trong Một Tháng",
        'compare_projects_year': "So Sánh Dự Án Trong Một Năm",
        'compare_projects_over_time': "So Sánh Nhiều Dự Án Qua Các Tháng/Năm",
        'filter_data_for_comparison': "Lọc dữ liệu để so sánh",
        'select_years': "Chọn năm(các năm):", 
        'select_months_comp': "Chọn tháng(các tháng):", 
        'select_projects_comp': "Chọn dự án(các dự án):", 
        'generate_comparison_report_btn': "🚀 Tạo báo cáo so sánh",
        'no_data_after_filter_comparison': "⚠️ {}",
        'latest_update_date': "Dữ liệu được cập nhật đến ngày",
        'data_filtered_success': "✅ Dữ liệu đã được lọc thành công cho so sánh.",
        'comparison_data_preview': "Xem trước dữ liệu so sánh",
        'generating_comparison_excel': "Đang tạo báo cáo Excel so sánh...",
        'comparison_excel_generated': "✅ Báo cáo Excel so sánh đã được tạo: {}",
        'download_comparison_excel': "📥 Tải báo cáo Excel so sánh",
        'generating_comparison_pdf': "Đang tạo báo cáo PDF so sánh...",
        'comparison_pdf_generated': "✅ Báo cáo PDF so sánh đã được tạo: {}",
        'download_comparison_pdf': "📥 Tải báo cáo PDF so sánh",
        'failed_to_generate_comparison_excel': "❌ Đã xảy ra lỗi khi tạo báo cáo Excel so sánh.",
        'failed_to_generate_comparison_pdf': "❌ Đã xảy ra lỗi khi tạo báo cáo PDF so sánh.",
        'raw_data_preview_header': "Dữ liệu đầu vào thô (100 hàng đầu)",
        'no_raw_data': "Không có dữ liệu thô được tải.",
        'no_year_in_data': "Không có năm nào trong dữ liệu để chọn.",
        'user_guide': "Hướng dẫn sử dụng",
        'export_options': "Tùy chọn xuất báo cáo",
        'export_excel_option': "Xuất ra Excel (.xlsx)",
        'export_pdf_option': "Xuất ra PDF (.pdf)",
        'report_button': "Tạo báo cáo",
        'no_data': "Không có dữ liệu sau khi lọc",
        'report_done': "Đã tạo báo cáo",
        'download_excel': "Tải Excel",
        'download_pdf': "Tải PDF",
        'warning_select_export_format': "Vui lòng chọn ít nhất một định dạng xuất báo cáo (Excel hoặc PDF).",
        'error_generating_report': "Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại.",
        'select_at_least_two_projects_warning': "Vui lòng chọn ít nhất hai dự án để so sánh.",
        'select_years_for_over_time_months': "Chọn năm (hoặc các năm) bạn muốn so sánh:",
        'select_months_for_single_year': "Chọn tháng(các tháng) trong năm đã chọn:",
        'comparison_over_years_note': "Lưu ý: Bạn đã chọn nhiều năm. Báo cáo sẽ so sánh dữ liệu của dự án qua các năm đã chọn. Lựa chọn tháng sẽ bị bỏ qua.",
        'comparison_over_months_note': "Lưu ý: Báo cáo sẽ so sánh dữ liệu của dự án qua các tháng đã chọn trong năm {}.",
        'no_comparison_criteria_selected': "Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.",
        'no_month_selected_for_single_year': "Vui lòng chọn ít nhất một tháng khi so sánh một dự án trong một năm cụ thể.",
        'tab_help': "Trợ giúp",
        'preview_charts_title': "📊 Biểu đồ xem trước",
        "help_instruction_simple": "Nếu bạn có bất kỳ thắc mắc nào hoặc cần hỗ trợ, vui lòng gửi email đến Quản trị viên **ky@triaccomposites.com**. Chúng tôi sẽ phản hồi trong thời gian sớm nhất. Xin cảm ơn!",
        'select_all_projects_checkbox': "Chọn tất cả dự án"
    }
}

def get_text(key, lang=None):
    lang = lang or st.session_state.get("lang", "vi")
    val = TEXTS.get(lang, {}).get(key)
    if val is None:
        return f"Missing text for {key}"
    if isinstance(val, tuple):
        return val[0] if lang == 'vi' else val[1]
    return val

col_logo_title, col_lang = st.columns([0.8, 0.2])
with col_logo_title:
    st.image("triac_logo.png", width=110)
    st.markdown("<div class='report-title'>Triac Time Report Generator</div>", unsafe_allow_html=True)
    st.markdown("<div class='report-subtitle'>Reporting tool for time tracking and analysis</div>", unsafe_allow_html=True)

with col_lang:
    selected_lang = st.radio(
        "Select language:",
        options=['vi', 'en'],
        format_func=lambda x: "Tiếng Việt" if x == "vi" else "English",
        key='language_selector_main'
    )
    if st.session_state.lang != selected_lang:
        st.session_state.lang = selected_lang

if not os.path.exists(path_dict['template_file']):
    st.error(get_text('template_not_found').format(path_dict['template_file']))
    st.stop()

@st.cache_data(ttl=1800)
def cached_load():
    df_raw = load_raw_data(path_dict['template_file'])
    df = df_raw.copy()
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Date'])

    df['Year'] = df['Date'].dt.year.astype(int)
    df['Month'] = df['Date'].dt.month.astype(int)
    df['Week'] = df['Date'].dt.isocalendar().week.astype(int)

    # Khai báo dữ liệu cuối tuần ngay từ khâu tiền xử lý
    df['DayOfWeek'] = df['Date'].dt.dayofweek 
    df['IsWeekend'] = df['DayOfWeek'].isin([5, 6])
    
    # 🌙 LOGIC GIỜ TĂNG CA BUỔI TỐI (TRÊN 8.5 GIỜ / NGÀY)
    GIO_CHUAN = 8.5
    df['Night_OT_Hours'] = (df['Hours'] - GIO_CHUAN).clip(lower=0)
    
    config_data = read_configs(path_dict['template_file'])
    return df, config_data

with st.spinner(get_text('loading_data')):
    df, config_data = cached_load()
    df_raw = df.copy()

if 'Date' in df_raw.columns:
    latest_date = pd.to_datetime(df_raw['Date'], errors='coerce').max()
    if pd.notnull(latest_date):
        st.info(f"📅 {get_text('latest_update_date')}: {latest_date.strftime('%d/%m/%Y')}")
    else:
        st.warning(get_text('no_valid_dates_found'))
else:
    st.warning(get_text('date_column_missing'))

if df_raw.empty:
    st.error(get_text('failed_to_load_raw_data'))
    st.stop()
    
def create_hierarchy_chart(df, level="Full"):
    level_options = {
        "Workcentre": ['Project name', 'Team', 'Workcentre'],
        "Task": ['Project name', 'Team', 'Workcentre', 'Task'],
        "Job": ['Project name', 'Team', 'Workcentre', 'Task', 'Job'],
        "Employee": ['Project name', 'Team', 'Workcentre', 'Task', 'Job', 'Employee'],
        "Full": ['Project name', 'Team', 'Workcentre', 'Task', 'Job', 'Employee']
    }
    if isinstance(level, dict):
        level = level.get("level", "Full")
    path_levels = level_options.get(str(level), level_options["Full"])
    required_cols = path_levels + ['Hours']
    if df.empty or not all(col in df.columns for col in required_cols):
        return None
    for col in path_levels:
        df[col] = df[col].fillna("Unknown")
    if 'Team leader' not in df.columns:
        df['Team leader'] = 'Unknown'
    fig = px.treemap(
        df,
        path=path_levels,
        values='Hours',
        hover_data=['Team leader'],
        title=f'📌 Hierarchical View: {" → ".join(path_levels)}',
        template='plotly_white'
    )
    return fig

all_years = sorted(df_raw['Year'].dropna().unique().astype(int).tolist())
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
all_months = [m for m in month_order if m in df_raw['MonthName'].dropna().unique()]
all_projects = sorted(df_raw['Project name'].dropna().unique().tolist())

tab_dashboard_main, tab_standard_report_main, tab_comparison_report_main, tab_data_preview_main, tab_user_guide_main, tab_help_main = st.tabs([
    "📈 Dashboard",
    get_text('tab_standard_report'),
    get_text('tab_comparison_report'),
    get_text('tab_data_preview'),
    get_text('user_guide'),
    get_text("tab_help")
])

def create_monthly_chart(df_filtered, config):
    if 'MonthName' not in df_filtered.columns or 'Hours' not in df_filtered.columns:
        return None
    ordered_months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]
    df_month = (
        df_filtered.groupby('MonthName')['Hours']
        .sum()
        .reindex(ordered_months)
        .dropna()
        .reset_index()
    )
    fig = px.bar(
        df_month,
        x='MonthName',
        y='Hours',
        title="📆 Monthly Total Hours",
        color='MonthName',
        template='plotly_white'
    )
    fig.update_layout(xaxis_title="Month", yaxis_title="Hours")
    return fig

def create_task_chart(df_filtered, config):
    if 'Task' not in df_filtered.columns or 'Hours' not in df_filtered.columns:
        return None
    df_task = (
        df_filtered.groupby('Task')['Hours']
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    fig = px.bar(
        df_task,
        x='Hours',
        y='Task',
        orientation='h',
        title="📋 Total Hours by Task",
        color='Task',
        template='plotly_white'
    )
    fig.update_layout(xaxis_title="Hours", yaxis_title="Task")
    return fig

def create_workcentre_chart(df_filtered, config):
    if 'Workcentre' not in df_filtered.columns or 'Hours' not in df_filtered.columns:
        return None
    df_wc = (
        df_filtered.groupby('Workcentre')['Hours']
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    fig = px.bar(
        df_wc,
        x='Hours',
        y='Workcentre',
        orientation='h',
        title="🏭 Total Hours by Workcentre",
        color='Workcentre',
        template='plotly_white'
    )
    fig.update_layout(xaxis_title="Hours", yaxis_title="Workcentre")
    return fig

def create_team_chart(df, config_data=None):
    if df.empty or not all(col in df.columns for col in ['Team', 'Team leader', 'Hours']):
        return None
    team_summary = (
        df.groupby(['Team', 'Team leader'])['Hours']
        .sum()
        .reset_index()
        .sort_values(by='Hours', ascending=False)
    )
    fig = px.bar(
        team_summary,
        x='Team',
        y='Hours',
        color='Team',
        hover_data=['Team leader'],
        title='👥 Total Hours by Team and Leader',
        template='plotly_white'
    )
    return fig

# =========================================================================
# HELPER FUNCTION: DRAW DYNAMIC OT VISUALS (Dùng chung cho cả 2 Dashboard)
# =========================================================================
def render_ot_dashboard_analytics(df_scope, is_project_level=False):
    """Hàm dựng các biểu đồ bóc tách và phân tích dữ liệu OT động"""
    df_scope = df_scope.copy()
    df_scope['Weekend_OT_Hours'] = df_scope.apply(lambda r: r['Hours'] if r['IsWeekend'] else 0.0, axis=1)
    
    st.markdown("#### 🔍 Bộ lọc đối tượng Tăng ca (OT Search Filter)")
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        wcs = sorted(df_scope['Workcentre'].dropna().unique().tolist())
        selected_wcs = st.multiselect("Lọc theo Nhóm (Workcentre):", wcs, default=wcs, key=f"wc_ot_f_{is_project_level}")
    with col_f2:
        tsks = sorted(df_scope['Task'].dropna().unique().tolist())
        selected_tsks = st.multiselect("Lọc theo Công việc (Task):", tsks, default=tsks, key=f"tsk_ot_f_{is_project_level}")
    with col_f3:
        emps = sorted(df_scope['Employee'].dropna().unique().tolist())
        selected_emps = st.multiselect("Lọc theo Nhân viên (Employee):", emps, default=emps, key=f"emp_ot_f_{is_project_level}")
        
    # Áp dụng bộ lọc
    df_filtered = df_scope[
        (df_scope['Workcentre'].isin(selected_wcs)) &
        (df_scope['Task'].isin(selected_tsks)) &
        (df_scope['Employee'].isin(selected_emps))
    ]
    
    if df_filtered.empty:
        st.info("ℹ️ Không tìm thấy bản ghi tăng ca nào khớp với bộ lọc đối tượng đã chọn.")
        return

    # Tabs phân rã chi tiết
    t_emp, t_tsk, t_wc = st.tabs(["👥 Bóc tách theo Nhân viên", "🛠️ Bóc tách theo Công việc", "🏭 Bóc tách theo Nhóm"])
    with t_emp:
        df_emp_ot = df_filtered.groupby('Employee')[['Weekend_OT_Hours', 'Night_OT_Hours']].sum().reset_index()
        df_emp_ot['Total_OT'] = df_emp_ot['Weekend_OT_Hours'] + df_emp_ot['Night_OT_Hours']
        df_emp_ot = df_emp_ot.sort_values(by='Total_OT', ascending=False)
        
        fig = px.bar(df_emp_ot.head(15), x='Employee', y=['Weekend_OT_Hours', 'Night_OT_Hours'], 
                     title="Top Nhân viên phát sinh Giờ OT nhiều nhất", text_auto='.1f')
        st.plotly_chart(fig, width='stretch')
        st.dataframe(df_emp_ot, width='stretch')
        
    with t_tsk:
        df_tsk_ot = df_filtered.groupby('Task')[['Weekend_OT_Hours', 'Night_OT_Hours']].sum().reset_index()
        df_tsk_ot['Total_OT'] = df_tsk_ot['Weekend_OT_Hours'] + df_tsk_ot['Night_OT_Hours']
        df_tsk_ot = df_tsk_ot.sort_values(by='Total_OT', ascending=False)
        
        fig = px.bar(df_tsk_ot.head(15), x='Task', y=['Weekend_OT_Hours', 'Night_OT_Hours'], 
                     title="Các Công việc (Task) phát sinh Giờ OT nhiều nhất", text_auto='.1f')
        st.plotly_chart(fig, width='stretch')
        st.dataframe(df_tsk_ot, width='stretch')
        
    with t_wc:
        df_wc_ot = df_filtered.groupby('Workcentre')[['Weekend_OT_Hours', 'Night_OT_Hours']].sum().reset_index()
        df_wc_ot['Total_OT'] = df_wc_ot['Weekend_OT_Hours'] + df_wc_ot['Night_OT_Hours']
        df_wc_ot = df_wc_ot.sort_values(by='Total_OT', ascending=False)
        
        fig = px.bar(df_wc_ot, x='Workcentre', y=['Weekend_OT_Hours', 'Night_OT_Hours'], 
                     title="Phân bổ Giờ OT theo các Nhóm (Workcentre)", text_auto='.1f')
        st.plotly_chart(fig, width='stretch')
        st.dataframe(df_wc_ot, width='stretch')


# =========================================================================
# STANDARD REPORT TAB
# =========================================================================
with tab_standard_report_main:
    st.header(get_text('standard_report_header'))
    col1_std, col2_std, col3_std = st.columns(3)
    with col1_std:
        if 'standard_analysis_mode' not in st.session_state:
            st.session_state.standard_analysis_mode = config_data['mode'] if config_data['mode'] in ['year', 'month', 'week'] else 'year'
        mode_options = ['year', 'month', 'week']
        try:
            mode_index = mode_options.index(st.session_state.standard_analysis_mode)
        except ValueError:
            mode_index = 0
            st.session_state.standard_analysis_mode = mode_options[0]
        mode = st.selectbox(
            get_text('select_analysis_mode'),
            options=mode_options,
            index=mode_index,
            key='standard_mode_tab'
        )
    with col2_std:
        if 'standard_selected_years' not in st.session_state:
            default_year = config_data['year'] if config_data['year'] in all_years else (all_years[0] if all_years else None)
            st.session_state.standard_selected_years = [default_year] if default_year else []
        selected_years = st.multiselect(
            get_text('select_year'),
            options=all_years,
            default=st.session_state.standard_selected_years,
            key='standard_year_tab'
        )
        if selected_years:
            st.session_state.standard_selected_years = selected_years
        else:
            st.warning(get_text('no_year_in_data'))
            st.stop()
    with col3_std:
        if 'standard_selected_months' not in st.session_state:
            st.session_state.standard_selected_months = config_data['months'] if config_data['months'] else all_months
        valid_default_months = [m for m in st.session_state.standard_selected_months if m in all_months]
        if not valid_default_months and all_months:
            valid_default_months = all_months
        selected_months = st.multiselect(
            get_text('select_months'),
            options=all_months,
            default=valid_default_months,
            key='standard_months_tab'
        )
        st.session_state.standard_selected_months = selected_months

    st.subheader(get_text('standard_project_selection_header'))
    initial_included_projects_config = config_data['project_filter_df'][
        config_data['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
    ]['Project Name'].tolist()

    if 'standard_selected_projects' not in st.session_state:
        default_standard_projects = [p for p in initial_included_projects_config if p in all_projects]
        if not default_standard_projects and all_projects:
            default_standard_projects = all_projects
        st.session_state.standard_selected_projects = default_standard_projects

    if "select_all_std_projects_checkbox" not in st.session_state:
        st.session_state.select_all_std_projects_checkbox = True

    select_all_std_projects = st.checkbox(
        get_text("select_all_projects_checkbox"), 
        key="select_all_std_projects_checkbox"
    )
    if select_all_std_projects:
        standard_project_selection = all_projects
    else:
        current_std_projects_default = [p for p in st.session_state.standard_selected_projects if p in all_projects]
        if not current_std_projects_default and all_projects:
            current_std_projects_default = all_projects
        st.caption(f"Đang chọn {len(current_std_projects_default)} dự án")
        standard_project_selection = st.multiselect(
            get_text('standard_project_selection_text'),
            options=all_projects,
            default=current_std_projects_default,
            key='standard_project_selection_tab'
        )
    if st.session_state.standard_selected_projects != standard_project_selection:
        st.session_state.standard_selected_projects = standard_project_selection
        
    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_std')
    export_pdf = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_std')

    if st.button(get_text('generate_standard_report_btn'), key='generate_standard_report_btn_tab'):
        if not export_excel and not export_pdf:
            st.warning(get_text("warning_select_export_format"))
        elif not selected_years:
            st.error(get_text('no_year_selected_error'))
        elif not standard_project_selection:
            st.warning(get_text('no_project_selected_warning_standard'))
        else:
            temp_project_filter_df_standard = pd.DataFrame({
                'Project Name': standard_project_selection,
                'Include': ['yes'] * len(standard_project_selection)
            })
            standard_report_config = {
                'mode': mode,
                'year': selected_years,
                'months': selected_months,
                'project_filter_df': temp_project_filter_df_standard
            }
            df_filtered_standard = apply_filters(df_raw, standard_report_config)
            project_col = 'Project name'
            valid_projects_in_filtered = df_filtered_standard[project_col].unique().tolist()
            standard_project_selection = [p for p in standard_project_selection if p in valid_projects_in_filtered]
            if not standard_project_selection:
                st.warning("Không có dự án nào có dữ liệu trong năm và tháng đã chọn.")
                st.stop()
            temp_project_filter_df_standard = pd.DataFrame({
                'Project Name': standard_project_selection,
                'Include': ['yes'] * len(standard_project_selection)
            })
            standard_report_config['project_filter_df'] = temp_project_filter_df_standard
            if 'Date' in df_filtered_standard.columns:
                df_filtered_standard['MonthName'] = pd.to_datetime(df_filtered_standard['Date']).dt.strftime('%B')
            if df_filtered_standard.empty:
                st.warning(get_text('no_data_after_filter_standard'))
            else:
                st.subheader(get_text("preview_charts_title"))
                fig_monthly = create_monthly_chart(df_filtered_standard, standard_report_config)
                if fig_monthly:
                    st.plotly_chart(fig_monthly, width='stretch')
                fig_task = create_task_chart(df_filtered_standard, standard_report_config)
                if fig_task:
                    st.plotly_chart(fig_task, width='stretch')
                fig_workcentre = create_workcentre_chart(df_filtered_standard, standard_report_config)
                if fig_workcentre:
                    st.plotly_chart(fig_workcentre, width='stretch')
                st.markdown("### 🧭 Chọn cấp độ phân tích")
                hierarchy_level = st.selectbox(
                    "Chọn cấp phân tích cho biểu đồ phân cấp:",
                    ["Workcentre", "Task", "Job", "Employee", "Full"],
                    index=4,
                    key="hierarchy_level_std"
                )
                fig_hierarchy = create_hierarchy_chart(df_filtered_standard, hierarchy_level)
                if fig_hierarchy:
                    st.plotly_chart(fig_hierarchy, width='stretch')
                    
                # 🌙 🔎 THÊM PHÂN ĐOẠN PHÂN TÍCH TĂNG CA (OT) CHO RIÊNG CÁC DỰ ÁN ĐƯỢC CHỌN TRONG BÁO CÁO TIÊU CHUẨN
                st.markdown("---")
                st.subheader("🌙 & 📅 Selected Projects OT Dashboard Preview")
                render_ot_dashboard_analytics(df_filtered_standard, is_project_level=True)
                
                st.markdown("---")
                today_str = datetime.today().strftime("%Y-%m-%d")
                path_dict = {
                    'output_file': f'outputs/standard/Time_report_Standard_{today_str}.xlsx',
                    'pdf_report': f'outputs/standard/Time_report_Standard_{today_str}.pdf',
                    'logo_path': 'triac_logo.png'
                } 
                report_generated = False
                
                # Khởi tạo thư mục trước khi lưu để tránh FileNotFoundError trên Cloud
                os.makedirs("outputs/standard", exist_ok=True)
                
                if export_excel:
                    with st.spinner(get_text('generating_excel_report')):
                        excel_success = export_report(df_filtered_standard, standard_report_config, path_dict['output_file'])
                    if excel_success:
                        st.success(get_text('excel_report_generated').format(os.path.basename(path_dict['output_file'])))
                        report_generated = True
                    else:
                        st.error(get_text('failed_to_generate_excel'))
                if export_pdf:
                    pdf_report_path = path_dict['pdf_report']
                    if not pdf_report_path:
                        raise ValueError("❌ pdf_report_path is empty. Please check where it's defined.")
                    with st.spinner(get_text('generating_pdf_report')):
                        pdf_success = export_pdf_report(df_filtered_standard, standard_report_config, path_dict['pdf_report'], path_dict['logo_path'])
                    if pdf_success:
                        st.success(get_text('pdf_report_generated').format(os.path.basename(path_dict['pdf_report'])))
                        report_generated = True
                    else:
                        st.error(get_text('failed_to_generate_pdf'))

                if report_generated:
                    if export_excel and os.path.exists(path_dict['output_file']):
                        with open(path_dict['output_file'], "rb") as f:
                            st.download_button(get_text("download_excel"), data=f, file_name=os.path.basename(path_dict['output_file']), width='stretch', key='download_excel_std_btn')
                    if export_pdf and os.path.exists(path_dict['pdf_report']):
                        with open(path_dict['pdf_report'], "rb") as f:
                            st.download_button(get_text("download_pdf"), data=f, file_name=os.path.basename(path_dict['pdf_report']), width='stretch', key='download_pdf_std_btn')
                else:
                    st.error(get_text('error_generating_report'))

# =========================================================================
# COMPARISON REPORT TAB
# =========================================================================
with tab_comparison_report_main:
    st.header(get_text('comparison_report_header'))
    internal_comparison_modes_map = {
        'compare_projects_month': ("So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"),
        'compare_projects_year': ("So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"),
        'compare_projects_over_time': ("So Sánh Nhiều Dự Án Qua Các Tháng/Năm", "Compare Projects Over Time (Months/Years)")
    }
    current_language = st.session_state.get("lang", "vi")
    comparison_mode_display_options = [
        vi if current_language == 'vi' else en
        for (_, (vi, en)) in internal_comparison_modes_map.items()
    ]
    display_to_internal_map = {
        (vi if current_language == 'vi' else en): key
        for key, (vi, en) in internal_comparison_modes_map.items()
    }
    default_key = st.session_state.get('selected_comparison_mode_key', list(internal_comparison_modes_map.keys())[0])
    vi_val, en_val = internal_comparison_modes_map[default_key]
    default_display = vi_val if current_language == 'vi' else en_val

    try:
        current_index = comparison_mode_display_options.index(default_display)
    except ValueError:
        current_index = 0
        default_key = list(internal_comparison_modes_map.keys())[0]
        st.session_state.selected_comparison_mode_key = default_key
        default_display = get_text(default_key)
        
    selected_display = st.selectbox(
        get_text('comparison_mode_label'),
        options=comparison_mode_display_options,
    )
    comparison_mode_selected = display_to_internal_map[selected_display]
    vi_val, en_val = internal_comparison_modes_map[comparison_mode_selected]
    comparison_mode = vi_val if st.session_state.lang == 'vi' else en_val
    
    st.subheader(get_text('filter_data_for_comparison'))
    if st.session_state.lang == 'vi':
        filter_mode_display_options = ["Theo Tổng Giờ", "Theo Task", "Theo Workcentre"]
    else:
        filter_mode_display_options = ["By Total hour", "By Task", "By Workcentre"]

    display_to_internal = {
        "Theo Tổng Giờ": "Total", "Theo Task": "Task", "Theo Workcentre": "Workcentre",
        "By Total hour": "Total", "By Task": "Task", "By Workcentre": "Workcentre"
    }
    current_display = st.session_state.get("selected_filter_display", filter_mode_display_options[0])
    if current_display not in filter_mode_display_options:
        current_display = filter_mode_display_options[0]
        st.session_state.selected_filter_display = current_display
        st.session_state.selected_filter_mode = display_to_internal[current_display]

    selected_filter_display = st.selectbox(
        "Comparison filter mode",
        options=filter_mode_display_options,
        index=filter_mode_display_options.index(current_display),
        key="filter_mode_selectbox"
    )
    if selected_filter_display != current_display:
        st.session_state.selected_filter_display = selected_filter_display
        st.session_state.selected_filter_mode = display_to_internal[selected_filter_display]

    filter_mode = st.session_state.get("selected_filter_mode", display_to_internal[current_display])

    if 'comparison_selected_projects' not in st.session_state:
        st.session_state.comparison_selected_projects = []
    validation_error = False

    if "select_all_projects_checkbox" not in st.session_state:
        st.session_state.select_all_projects_checkbox = True

    select_all_projects = st.checkbox(
        get_text("select_all_projects_checkbox"),
        key="select_all_projects_checkbox"
    )
    if select_all_projects:
        comp_projects = all_projects
        if st.session_state.comparison_selected_projects != all_projects:
            st.session_state.comparison_selected_projects = all_projects
    else:
        comp_projects = st.multiselect(
            get_text('select_projects_comp'),
            options=all_projects,
            default=st.session_state.comparison_selected_projects,
            key='comp_projects_select_tab_common'
        )
        if comp_projects != st.session_state.comparison_selected_projects:
            st.session_state.comparison_selected_projects = comp_projects

    if comparison_mode == "So Sánh Nhiều Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare Projects Over Time (Months/Years)":
        if len(comp_projects) < 1:
            st.warning(get_text('no_project_selected_warning_standard'))
            validation_error = True
        if 'comparison_selected_years_over_time' not in st.session_state:
            st.session_state.comparison_selected_years_over_time = []
        selected_years_over_time = st.multiselect(
            get_text('select_years_for_over_time_months'),
            options=all_years,
            default=st.session_state.comparison_selected_years_over_time,
            key='comp_years_select_tab_over_time'
        )
        if selected_years_over_time != st.session_state.comparison_selected_years_over_time:
            st.session_state.comparison_selected_years_over_time = selected_years_over_time
        comp_years = selected_years_over_time

        if 'comparison_selected_months_over_time' not in st.session_state:
            st.session_state.comparison_selected_months_over_time = []

        if len(selected_years_over_time) == 1:
            st.info(get_text('comparison_over_months_note').format(selected_years_over_time[0]))
            comp_months = st.multiselect(
                get_text('select_months_for_single_year'),
                options=all_months,
                default=[m for m in st.session_state.comparison_selected_months_over_time if m in all_months],
                key='comp_months_select_tab_over_time'
            )
            st.session_state.comparison_selected_months_over_time = comp_months
            if not comp_months:
                st.warning(get_text('no_month_selected_for_single_year'))
                validation_error = True
        elif len(selected_years_over_time) > 1:
            st.info(get_text('comparison_over_years_note'))
            comp_months = []
            st.session_state.comparison_selected_months_over_time = []
        else:
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True
            comp_months = []
            st.session_state.comparison_selected_months_over_time = []

    elif comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month", "So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        col_comp1, col_comp2 = st.columns(2)
        with col_comp1:
            if 'comparison_selected_years_general' not in st.session_state:
                st.session_state.comparison_selected_years_general = []
            comp_years = st.multiselect(
                get_text('select_years'),
                options=all_years,
                default=[y for y in st.session_state.comparison_selected_years_general if y in all_years],
                key='comp_years_select_tab_general'
            )
            st.session_state.comparison_selected_years_general = comp_years
        with col_comp2:
            if 'comparison_selected_months_general' not in st.session_state:
                st.session_state.comparison_selected_months_general = []
            if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
                comp_months = st.multiselect(
                    get_text('select_months_comp'),
                    options=all_months,
                    default=[m for m in st.session_state.comparison_selected_months_general if m in all_months],
                    key='comp_months_select_tab_general'
                )
                st.session_state.comparison_selected_months_general = comp_months
            else:
                comp_months = []
                st.session_state.comparison_selected_months_general = []

        if not comp_years:
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True
        if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"] and not comp_months:
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True
        if not comp_projects:
            st.warning(get_text('no_project_selected_warning_standard'))
            validation_error = True

    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel_comp = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_comp')
    export_pdf_comp = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_comp')

    if st.button(get_text('generate_comparison_report_btn'), key='generate_comparison_report_btn_tab'):
        if not export_excel_comp and not export_pdf_comp:
            st.warning(get_text("warning_select_export_format"))
        elif validation_error:
            pass
        else:
            comparison_config = {
                'years': comp_years,
                'months': comp_months,
                'selected_projects': comp_projects,
                'filter_mode': filter_mode
            }
            comparison_output_folder = "outputs/comparison"
            comparison_path_dict = path_dict.copy()
            comparison_path_dict.update({
                "comparison_output_excel": os.path.join(comparison_output_folder, "comparison_result.xlsx"),
                "comparison_output_file": os.path.join(comparison_output_folder, "comparison_export.xlsx"),
                "comparison_pdf_output": os.path.join(comparison_output_folder, "comparison_chart.png"),
                "comparison_pdf_report": os.path.join(comparison_output_folder, "comparison_report.pdf"),
                "logo": path_dict["logo_path"]
            })
            df_filtered_comparison, comparison_filter_message, filtered_projects = apply_comparison_filters(
                df_raw, comparison_config, comparison_mode, filter_mode
            )
            original_projects = comparison_config.get("selected_projects", [])
            if len(filtered_projects) < len(original_projects):
                removed = set(original_projects) - set(filtered_projects)
                st.warning(f"⚠️ Một số dự án không có dữ liệu thực tế và đã bị loại khỏi báo cáo: {', '.join(removed)}")

            if df_filtered_comparison.empty:
                for key in ["comparison_output_excel", "comparison_pdf_output", "comparison_output_file", "comparison_pdf_report"]:
                    folder = os.path.dirname(comparison_path_dict[key])
                    if folder: os.makedirs(folder, exist_ok=True)
                st.warning(get_text('no_data_after_filter_comparison').format(comparison_filter_message))
            else:
                st.success(get_text('data_filtered_success'))
                st.subheader(get_text('comparison_data_preview'))
                st.dataframe(df_filtered_comparison)
                st.subheader(get_text("preview_charts_title"))

                fig_monthly = create_monthly_chart(df_filtered_comparison, comparison_config)
                if fig_monthly: st.plotly_chart(fig_monthly, width='stretch')

                fig_task = create_task_chart(df_filtered_comparison, comparison_config)
                if fig_task: st.plotly_chart(fig_task, width='stretch')

                fig_workcentre = create_workcentre_chart(df_filtered_comparison, comparison_config)
                if fig_workcentre: st.plotly_chart(fig_workcentre, width='stretch')
                    
                if 'df_filtered_comparison' in locals():
                    fig_hierarchy = create_hierarchy_chart(df_filtered_comparison, comparison_config)
                    if fig_hierarchy: st.plotly_chart(fig_hierarchy, width='stretch')
                st.markdown("---")

                report_generated_comp = False
                
                # Khởi tạo thư mục trước khi lưu để tránh FileNotFoundError trên Cloud
                os.makedirs("outputs/comparison", exist_ok=True)
                
                if export_excel_comp:
                    with st.spinner(get_text('generating_comparison_excel')):
                        try:
                            excel_success_comp = export_comparison_report(
                                df_filtered_comparison, comparison_config,
                                comparison_path_dict['comparison_output_file'],
                                comparison_mode, filter_mode
                            )
                        except Exception as e:
                            excel_success_comp = False
                            st.error(f"❌ Lỗi khi xuất Excel: {e}")
                    if os.path.exists(comparison_path_dict['comparison_output_file']):
                        st.success("✅ File Excel đã được tạo đúng tại: " + comparison_path_dict['comparison_output_file'])
                        report_generated_comp = True
                    else:
                        st.error("❌ File Excel KHÔNG được tạo ra: " + comparison_path_dict['comparison_output_file'])
                    if excel_success_comp:
                        st.success(get_text('comparison_excel_generated').format(os.path.basename(comparison_path_dict['comparison_output_file'])))
                        report_generated_comp = True
                    else:
                        st.error(get_text('failed_to_generate_comparison_excel'))

                if export_pdf_comp:
                    with st.spinner(get_text('generating_comparison_pdf')):
                        try:
                            pdf_path = comparison_path_dict['comparison_pdf_report']
                            pdf_success_comp = export_comparison_pdf_report(
                                df_filtered_comparison, comparison_config,
                                pdf_path, comparison_mode,
                                comparison_path_dict['logo'], filter_mode
                            )
                        except Exception as e:
                            pdf_success_comp = False
                            st.error(f"❌ Lỗi khi xuất PDF: {e}")
                    if pdf_success_comp:
                        st.success(get_text('comparison_pdf_generated').format(os.path.basename(comparison_path_dict['comparison_pdf_report'])))
                        report_generated_comp = True
                    else:
                        st.error(get_text('failed_to_generate_comparison_pdf'))
                
                if report_generated_comp:
                    with st.expander("📥 Tải báo cáo PDF/Excel so sánh"):
                        excel_path = comparison_path_dict.get("comparison_output_file")
                        pdf_path = comparison_path_dict.get("comparison_pdf_report")
                        if export_excel_comp and excel_path and os.path.exists(excel_path):
                            with open(excel_path, "rb") as f_excel:
                                excel_data = f_excel.read()
                            st.download_button(
                                label="📄 Tải Excel So sánh", data=excel_data,
                                file_name=os.path.basename(comparison_path_dict["comparison_output_file"]),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width='stretch', key="exp_excel_comp_btn"
                            )
                        if export_pdf_comp and pdf_path and os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as f_pdf:
                                pdf_data = f_pdf.read()
                            st.download_button(
                                label="🖨️ Tải PDF So sánh", data=pdf_data,
                                file_name=os.path.basename(comparison_path_dict["comparison_pdf_report"]),
                                mime="application/pdf", width='stretch', key="exp_pdf_comp_btn"
                            )
                else:
                    st.error(get_text("⚠️ error_generating_report"))

# =========================================================================
# DATA PREVIEW TAB
# =========================================================================
with tab_data_preview_main:
    st.subheader(get_text('raw_data_preview_header'))
    if not df_raw.empty:
        st.dataframe(df_raw.head(100))
    else:
        st.info(get_text('no_raw_data'))

# =========================================================================
# USER GUIDE TAB
# =========================================================================
with tab_user_guide_main:
    st.markdown(f"### {get_text('user_guide')}")
    st.markdown("""
    - Select filters: mode, year, month, project
    - Select report export format (Excel, PDF or both)
    - Click "Create report"
    - Download generated report
    """)
    if "access_log" in st.session_state:
        st.write("📜 Current session access log:")
        st.dataframe(pd.DataFrame(st.session_state.access_log))

# =========================================================================
# HELP TAB
# =========================================================================
with tab_help_main:
    lang = st.session_state.get("lang", "vi")
    st.markdown(f"### {get_text('tab_help', lang)}")
    st.markdown(get_text("help_instruction_simple", lang))

# =========================================================================
# DASHBOARD TAB (ĐÃ ĐƯỢC TÍCH HỢP TOÀN DIỆN THÊM WEEKEND & NIGHT OT THEO TUẦN/THÁNG)
# =========================================================================
with tab_dashboard_main:
    template_name = "plotly_white" if "plotly_white" in pio.templates else None
    st.subheader("📊 Quick Overview Dashboard")

    available_years = sorted(df['Year'].dropna().unique().astype(int))
    if not available_years:
        st.error("❌ No year data available")
        st.stop()

    current_year = st.selectbox(
        "📅 Select year",
        available_years,
        index=len(available_years) - 1
    )

    if df['Month'].dtype == 'O':
        month_str_to_num = {
            month: i for i, month in enumerate(
                [datetime(1900, m, 1).strftime('%B') for m in range(1, 13)], start=1
            )
        }
        df['Month'] = df['Month'].map(month_str_to_num)

    available_months = sorted(
        pd.to_numeric(df[df['Year'] == current_year]['Month'], errors='coerce')
        .dropna().astype(int).loc[lambda x: (x >= 1) & (x <= 12)].unique()
    )
    if not available_months:
        st.warning(f"No month data available for year {current_year}")
        st.stop()

    month_name_map = {i: datetime(1900, i, 1).strftime('%B') for i in available_months}
    month_options = {f"{month_name_map[m]} {current_year}": (current_year, m) for m in available_months}

    selected_month_label = st.selectbox("📅 Select month", list(month_options.keys()))
    if selected_month_label not in month_options:
        st.stop()

    current_year, current_month = month_options[selected_month_label]
    current_month_name = month_name_map[current_month]

    def get_week_date_range(year, week_num):
        try:
            d = datetime.strptime(f'{year}-W{int(week_num)}-1', "%Y-W%W-%w")
            start_date = d.strftime('%d/%m')
            end_date = (d + timedelta(days=6)).strftime('%d/%m')
            return f"Week {week_num} ({start_date} → {end_date})"
        except Exception:
            return f"Week {week_num}"

    df_month = df[(df['Year'] == current_year) & (df['Month'] == current_month)]
    available_weeks = sorted(df_month['Week'].dropna().unique())

    if available_weeks:
        week_labels = {w: get_week_date_range(current_year, int(w)) for w in available_weeks}
        selected_week_num = st.selectbox(
            "🗓️ Select a week in the selected month (optional)",
            options=[None] + list(available_weeks),
            format_func=lambda x: week_labels.get(x, f"Week {x}") if x is not None else "📅 All Weeks in Month",
            index=0
        )
        df_week = df_month if selected_week_num is None else df_month[df_month['Week'] == selected_week_num]
    else:
        st.warning("⚠️ No weekly data found for selected month.")
        df_week = df_month
        selected_week_num = None

    # --- TÍNH TOÁN CÁC BIẾN SỐ KPI ---
    total_hours_week = df_week['Hours'].sum()
    total_hours_month = df_month['Hours'].sum()
    
    # 🚨 Tính giờ OT dựa theo phạm vi lọc hiện thời (Theo Tuần / Hoặc cả Tháng)
    total_weekend_hours_filtered = df_week[df_week['IsWeekend'] == True]['Hours'].sum()
    total_night_hours_filtered = df_week['Night_OT_Hours'].sum()

    # 🛠️ Giao diện KPI 4 cột trực quan
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        st.metric("🗓️ Total Scope Hours", f"{total_hours_week:.1f}h")
    with m_col2:
        st.metric("📆 Total Monthly Hours", f"{total_hours_month:.1f}h")
    with m_col3:
        st.metric("📅 Total Weekend OT Hours", f"{total_weekend_hours_filtered:.1f}h")
    with m_col4:
        st.metric("🌙 Total Night OT Hours", f"{total_night_hours_filtered:.1f}h")

    # =========================================================================
    # Phân đoạn Dashboard OT nâng cao: Hỗ trợ truy xuất sâu đối tượng
    # =========================================================================
    st.markdown("---")
    st.subheader("🚨 Overtime Analytics (Phân tích chi tiết Tăng ca Đêm & Cuối tuần)")
    render_ot_dashboard_analytics(df_week, is_project_level=False)

    # =========================================================================
    # GIỮ NGUYÊN CÁC BIỂU ĐỒ TỔNG QUAN KHÁC CỦA BẠN PHÍA DƯỚI
    # =========================================================================
    st.markdown("---")
    st.subheader("🔝 Top 5 Projects")
    top_projects = (
        df_week.groupby("Project name")["Hours"]
        .sum()
        .sort_values(ascending=False)
        .head(5)
        .reset_index()
    )
    fig1 = px.bar(
        top_projects, x="Project name", y="Hours", color="Project name",
        title="🔝 Top 5 Projects by Hours", template=template_name
    )
    st.plotly_chart(fig1, width='stretch')

    if all(col in df_week.columns for col in ["Workcentre", "Team leader"]):
        team_leader_ratio = (
            df_week.groupby(["Workcentre", "Team leader"])["Hours"]
            .sum()
            .reset_index()
        )
        fig2 = px.pie(
            team_leader_ratio,
            names="Workcentre",
            values="Hours",
            hover_data=["Team leader"],
            title="🧩 Hour Distribution by Team",
            template=template_name
        )
    else:
        team_ratio = df_week.groupby("Workcentre")["Hours"].sum().reset_index()
        fig2 = px.pie(
            team_ratio,
            names="Workcentre", values="Hours",
            title="🧩 Hour Distribution by Team", template=template_name
        )
    st.plotly_chart(fig2, width='stretch')

    if all(col in df_week.columns for col in ["Project name", "Workcentre", "Team leader"]):
        team_project = (
            df_week.groupby(["Project name", "Workcentre", "Team leader"])["Hours"]
            .sum()
            .reset_index()
        )
        fig3 = px.bar(
            team_project,
            x="Project name",
            y="Hours",
            color="Workcentre",
            hover_data=["Team leader"],
            title="🏗️ Team Allocation by Project",
            template=template_name
        )
    else:
        team_project = df_week.groupby(["Project name", "Workcentre"])["Hours"].sum().reset_index()
        fig3 = px.bar(
            team_project,
            x="Project name",
            y="Hours",
            color="Workcentre",
            title="🏗️ Team Allocation by Project",
            template=template_name
        )
    st.plotly_chart(fig3, width='stretch')

    if all(col in df_week.columns for col in ['Team', 'Team leader', 'Employee']):
        st.subheader("👥 Total Hours by Team, Leader and Employee")
        df_team_emp = (
            df_week.groupby(['Team', 'Team leader', 'Employee'])['Hours']
            .sum()
            .reset_index()
        )
        fig_team_emp = px.bar(
            df_team_emp,
            x="Team",
            y="Hours",
            color="Employee",
            hover_data=["Team leader"],
            title="👥 Total Hours by Team and Employee",
            template=template_name
        )
        fig_team_emp.update_layout(barmode='stack', xaxis_title="Team", yaxis_title="Total Hours")
        st.plotly_chart(fig_team_emp, width='stretch')
    else:
        st.info("⚠️ Not enough data to display team + employee breakdown.")

    st.markdown("---")
    st.subheader("🧭 Hierarchical Analysis (Project → Team → Workcentre → Task → Job → Employee)")
    df_hierarchy_base = df_week if not df_week.empty else df_month
    required_cols = ['Project name', 'Team', 'Workcentre', 'Task', 'Job', 'Employee', 'Hours']

    if all(col in df_hierarchy_base.columns for col in required_cols):
        fig_hierarchy = create_hierarchy_chart(df_hierarchy_base)
        if fig_hierarchy:
            st.plotly_chart(fig_hierarchy, width='stretch')
        else:
            st.info("⚠️ Not enough data to generate the hierarchy chart.")
    else:
        st.info("⚠️ Not enough data to display hierarchy chart (columns required: Project name, Team, Workcentre, Task, Job, Employee, Hours)")
