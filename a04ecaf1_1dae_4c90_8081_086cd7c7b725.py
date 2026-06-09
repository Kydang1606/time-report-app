import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import tempfile
import re
import shutil
from pandas import Series
import traceback
import numpy as np
import smtplib
from email.mime.text import MIMEText
from collections import defaultdict

# Hàm hỗ trợ làm sạch tên file/sheet
def sanitize_filename(name):
    # Ký tự không hợp lệ trong tên file/sheet của Excel
    invalid_chars = re.compile(r'[\\/*?[\]:;|=,<>]')
    s = invalid_chars.sub("_", str(name))
    # Loại bỏ các ký tự điều khiển ASCII và các ký tự không an toàn khác
    s = ''.join(c for c in s if c.isprintable())
    return s[:31] # Giới hạn 31 ký tự cho tên sheet trong Excel

def setup_paths():
    """Thiết lập các đường dẫn file đầu vào và đầu ra."""
    today = datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm",
        'output_file': f"Time_report_Standard_{today}.xlsx",
        'pdf_report': f"Time_report_Standard_{today}.pdf",
        'comparison_output_file': f"Time_report_Comparison_{today}.xlsx",
        'comparison_pdf_report': f"Time_report_Comparison_{today}.pdf",
        'logo_path': "triac_logo.png" # Thêm đường dẫn logo
    }

def get_comparison_pdf_path(comparison_mode, base_path):
    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        return base_path.replace(".pdf", "_Month.pdf")
    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        return base_path.replace(".pdf", "_SingleProjMonths.pdf")
    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Năm", "Compare One Project Over Years"]:
        return base_path.replace(".pdf", "_SingleProjYears.pdf")
    else:
        return base_path
        
def get_comparison_excel_path(comparison_mode, base_path):
    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        return base_path.replace(".xlsx", "_Month.xlsx")
    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        return base_path.replace(".xlsx", "_SingleProjMonths.xlsx")
    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Năm", "Compare One Project Over Years"]:
        return base_path.replace(".xlsx", "_SingleProjYears.xlsx")
    else:
        return base_path

def read_configs(template_file):
    """Đọc cấu hình từ file template Excel."""
    try:
        year_mode_df = pd.read_excel(template_file, sheet_name='Config_Year_Mode', engine='openpyxl')
        project_filter_df = pd.read_excel(template_file, sheet_name='Config_Project_Filter', engine='openpyxl')

        # Xử lý mode, year, months an toàn hơn
        mode_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value']
        mode = str(mode_row.values[0]).strip().lower() if not mode_row.empty and pd.notna(mode_row.values[0]) else 'year'

        year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
        year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) and pd.api.types.is_number(year_row.values[0]) else datetime.now().year
        
        months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
        months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty Glen else []
        
        if 'Include' in project_filter_df.columns:
            project_filter_df['Include'] = project_filter_df['Include'].astype(str).str.lower()

        return {
            'mode': mode,
            'year': year,
            'months': months,
            'project_filter_df': project_filter_df
        }
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file template tại {template_file}")
        return {'mode': 'year', 'year': datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}
    except Exception as e:
        print(f"Lỗi khi đọc cấu hình: {e}")
        return {'mode': 'year', 'year': datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}

def load_raw_data(template_file):
    """Tải dữ liệu thô từ file template Excel và thiết lập thuộc tính ngày cuối tuần, tăng ca đêm."""
    try:
        df = pd.read_excel(template_file, sheet_name='Raw Data', engine='openpyxl')
        df.columns = df.columns.str.strip()
        df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee', 'Project Name': 'Project name'}, inplace=True)
        
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date']) 
        
        df['Year'] = df['Date'].dt.year
        df['MonthName'] = df['Date'].dt.month_name()
        df['Week'] = df['Date'].dt.isocalendar().week.astype(int)
        
        # Bộ phân loại cuối tuần
        df['DayOfWeek'] = df['Date'].dt.dayofweek 
        df['IsWeekend'] = df['DayOfWeek'].isin([5, 6])
        
        df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce').fillna(0)
        
        # Logic tính giờ làm đêm và giờ hành chính chuẩn
        GIO_CHUAN = 8.5
        df['Night_OT_Hours'] = (df['Hours'] - GIO_CHUAN).clip(lower=0)
        df['Normal_Hours'] = df['Hours'].clip(upper=GIO_CHUAN)
        
        return df
    except Exception as e:
        print(f"Lỗi khi tải dữ liệu thô: {e}")
        return pd.DataFrame()

def apply_filters(df, config):
    """Áp dụng các bộ lọc dữ liệu dựa trên cấu hình."""
    df_filtered = df.copy()

    if 'years' in config and config['years']:  
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']:  
        if isinstance(config['year'], list):
            df_filtered = df_filtered[df_filtered['Year'].isin(config['year'])]
        else:
            df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    if config.get('months'):
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    if not config['project_filter_df'].empty:
        selected_project_names = config['project_filter_df']['Project Name'].tolist()
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_project_names)]
    else:
        return pd.DataFrame(columns=df.columns)  

    return df_filtered

def export_report(df, config, output_file_path):
    """Xuất báo cáo tiêu chuẩn ra file Excel kèm theo phân tích Tăng ca cuối tuần & Tăng ca đêm."""
    mode = config.get('mode', 'year')
    
    groupby_cols = []
    if mode == 'year':
        groupby_cols = ['Year', 'Project name']
    elif mode == 'month':
        groupby_cols = ['Year', 'MonthName', 'Project name']
    else: 
        groupby_cols = ['Year', 'Week', 'Project name']

    for col in groupby_cols + ['Hours']:
        if col not in df.columns:
            print(f"Lỗi: Cột '{col}' không tồn tại.")
            return False

    if df.empty:
        return False

    try:
        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='RawData', index=False)

        wb = load_workbook(output_file_path)

        # Summary sheet
        summary_chart = df.groupby('MonthName')['Hours'].sum().reset_index()
        summary_chart = summary_chart.sort_values('MonthName', key=lambda x: pd.to_datetime(x, format='%B'))

        if 'Summary' in wb.sheetnames:
            ws = wb['Summary']
            wb.remove(ws)
        ws = wb.create_sheet("Summary", 0)

        ws.append(['MonthName', 'Hours'])
        for row in summary_chart.itertuples(index=False):
            ws.append([row[0], row[1]])

        data_ref = Reference(ws, min_col=2, min_row=1, max_row=1 + len(summary_chart))
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=1 + len(summary_chart))

        chart = BarChart()
        chart.title = "Total Hours by Month"
        chart.x_axis.title = "Month"
        chart.y_axis.title = "Hours"
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, "E2")

        # 📊 WEEKEND OT SUMMARY (Chỉ dựng khi thực sự có bản ghi phát sinh)
        df_weekend = df[df['IsWeekend'] == True]
        if not df_weekend.empty:
            ot_summary = df_weekend.groupby(['Project name', 'Task'])['Hours'].sum().reset_index().sort_values(by='Hours', ascending=False)
            if not ot_summary.empty:
                ws_ot = wb.create_sheet("Weekend OT Summary", 1)
                ws_ot.append(['Project name', 'Task', 'Weekend OT Hours'])
                for row in ot_summary.itertuples(index=False):
                    ws_ot.append([row[0], row[1], row[2]])
                
                proj_ot_chart_data = df_weekend.groupby('Project name')['Hours'].sum().reset_index()
                if not proj_ot_chart_data.empty:
                    ws_ot.append([])
                    ws_ot.append(['📊 Project Total Summary'])
                    ws_ot.append(['Project name', 'Total OT Hours'])
                    ot_start_row = ws_ot.max_row
                    for row in proj_ot_chart_data.itertuples(index=False):
                        ws_ot.append([row[0], row[1]])
                    ot_end_row = ws_ot.max_row
                    
                    chart_ot = BarChart()
                    chart_ot.title = "Weekend Overtime Hours by Project"
                    chart_ot.x_axis.title = "Project"
                    chart_ot.y_axis.title = "Hours"
                    data_ref_ot = Reference(ws_ot, min_col=2, min_row=ot_start_row, max_row=ot_end_row)
                    cats_ref_ot = Reference(ws_ot, min_col=1, min_row=ot_start_row+1, max_row=ot_end_row)
                    chart_ot.add_data(data_ref_ot, titles_from_data=True)
                    chart_ot.set_categories(cats_ref_ot)
                    ws_ot.add_chart(chart_ot, "E2")

        # 🌙 NIGHT OT SUMMARY (Chỉ dựng khi có người làm trên 8.5 tiếng)
        df_night_ot = df[df['Night_OT_Hours'] > 0]
        if not df_night_ot.empty:
            night_ot_summary = df_night_ot.groupby(['Project name', 'Employee', 'Task'])['Night_OT_Hours'].sum().reset_index().sort_values(by='Night_OT_Hours', ascending=False)
            if not night_ot_summary.empty:
                ws_night_ot = wb.create_sheet("Night OT Summary", 2)
                ws_night_ot.append(['Project name', 'Employee', 'Task', 'Night OT Hours'])
                for row in night_ot_summary.itertuples(index=False):
                    ws_night_ot.append([row[0], row[1], row[2], row[3]])
                    
                proj_night_chart_data = df_night_ot.groupby('Project name')['Night_OT_Hours'].sum().reset_index()
                if not proj_night_chart_data.empty:
                    ws_night_ot.append([])
                    ws_night_ot.append(['🌙 Project Night OT Summary'])
                    ws_night_ot.append(['Project name', 'Total Night OT Hours'])
                    n_start_row = ws_night_ot.max_row
                    for row in proj_night_chart_data.itertuples(index=False):
                        ws_night_ot.append([row[0], row[1]])
                    n_end_row = ws_night_ot.max_row
                    
                    chart_night = BarChart()
                    chart_night.title = "Night Overtime Hours by Project (>8.5h/day)"
                    chart_night.x_axis.title = "Project"
                    chart_night.y_axis.title = "Hours"
                    data_ref_n = Reference(ws_night_ot, min_col=2, min_row=n_start_row, max_row=n_end_row)
                    cats_ref_n = Reference(ws_night_ot, min_col=1, min_row=n_start_row+1, max_row=n_end_row)
                    chart_night.add_data(data_ref_n, titles_from_data=True)
                    chart_night.set_categories(cats_ref_n)
                    ws_night_ot.add_chart(chart_night, "F2")

        # Xử lý các sheet Project riêng lẻ
        for project in df['Project name'].unique():
            df_proj = df[df['Project name'] == project]
            sheet_title = sanitize_filename(project)
            ws_proj = wb[sheet_title] if sheet_title in wb.sheetnames else wb.create_sheet(title=sheet_title)

            summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index().sort_values('Hours', ascending=False)
            if not summary_task.empty:
                ws_proj.append(['Task', 'Hours'])
                for row_data in dataframe_to_rows(summary_task, index=False, header=False):
                    ws_proj.append(row_data)

                chart_task = BarChart()
                chart_task.title = f"{project} - Hours by Task"
                chart_task.x_axis.title = "Task"
                chart_task.y_axis.title = "Hours"
                data_ref_task = Reference(ws_proj, min_col=2, min_row=1, max_row=len(summary_task) + 1)
                cats_ref_task = Reference(ws_proj, min_col=1, min_row=2, max_row=len(summary_task) + 1)
                chart_task.add_data(data_ref_task, titles_from_data=True)
                chart_task.set_categories(cats_ref_task)
                ws_proj.add_chart(chart_task, "E1")

            df_proj_weekend = df_proj[df_proj['IsWeekend'] == True]
            if not df_proj_weekend.empty:
                summary_weekend_month = df_proj_weekend.groupby('MonthName')['Hours'].sum().reset_index()
                month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                summary_weekend_month['MonthName'] = pd.Categorical(summary_weekend_month['MonthName'], categories=month_order, ordered=True)
                summary_weekend_month = summary_weekend_month.sort_values('MonthName').dropna()

                if not summary_weekend_month.empty:
                    ws_proj.append([])
                    ws_proj.append([])
                    start_row_ot = ws_proj.max_row + 1
                    ws_proj.append(['Weekend OT Month', 'OT Hours'])
                    for row_data in dataframe_to_rows(summary_weekend_month, index=False, header=False):
                        ws_proj.append(row_data)
                    end_row_ot = ws_proj.max_row

                    chart_ot_month = LineChart()  
                    chart_ot_month.title = f"{project} - Weekend OT Trend by Month"
                    chart_ot_month.x_axis.title = "Month"
                    chart_ot_month.y_axis.title = "Hours"
                    data_ref_ot = Reference(ws_proj, min_col=2, min_row=start_row_ot, max_row=end_row_ot)
                    cats_ref_ot = Reference(ws_proj, min_col=1, min_row=start_row_ot + 1, max_row=end_row_ot)
                    chart_ot_month.add_data(data_ref_ot, titles_from_data=True)
                    chart_ot_month.set_categories(cats_ref_ot)
                    ws_proj.add_chart(chart_ot_month, "E16")

            start_row_raw_data = ws_proj.max_row + 2
            if not summary_task.empty: start_row_raw_data += 15 

            for r_idx, r in enumerate(dataframe_to_rows(df_proj, index=False, header=True)):
                for c_idx, cell_val in enumerate(r):
                    ws_proj.cell(row=start_row_raw_data + r_idx, column=c_idx + 1, value=cell_val)
        
        ws_config = wb.create_sheet("Config_Info")
        ws_config['A1'], ws_config['B1'] = "Mode", config.get('mode', 'N/A').capitalize()
        ws_config['A2'], ws_config['B2'] = "Year(s)", ', '.join(map(str, config.get('years', []))) if config.get('years') else str(config.get('year', 'N/A'))
        ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config.get('months', [])) if config.get('months') else "All"
        
        if 'project_filter_df' in config and not config['project_filter_df'].empty:
            selected_projects_display = config['project_filter_df'][config['project_filter_df']['Include'].astype(str).str.lower() == 'yes']['Project Name'].tolist()
            ws_config['A4'], ws_config['B4'] = "Projects Included", ', '.join(selected_projects_display)

        for sheet_name in ['Raw Data', 'Config_Year_Mode', 'Config_Project_Filter']:
            if sheet_name in wb.sheetnames: del wb[sheet_name]

        wb.save(output_file_path)
        return True
    except Exception as e:
        print(f"❌ Lỗi trong export_report: {e}")
        traceback.print_exc()
        return False

def export_pdf_report(df, config, pdf_report_path, logo_path):
    if not pdf_report_path: return False
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    try:
        summary_chart = df.groupby('MonthName')['Hours'].sum().reset_index()
        summary_chart = summary_chart.sort_values('MonthName', key=lambda x: pd.to_datetime(x, format='%B'))

        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.bar(summary_chart['MonthName'], summary_chart['Hours'], color='skyblue')
        ax.set_title("Tổng giờ theo tháng")
        ax.set_xlabel("Tháng")
        ax.set_ylabel("Giờ")
        ax.bar_label(bars, labels=[f"{v:.1f}" for v in summary_chart['Hours']], padding=3)
        plt.xticks(rotation=45)
        plt.tight_layout()
        chart_path = os.path.join(tmp_dir, "standard_month_chart.png")
        fig.savefig(chart_path, dpi=150)
        plt.close(fig)
        charts_for_pdf.append((chart_path, "Total hour by month", None))
        
        if 'Project name' in df.columns:
            for project in df['Project name'].dropna().unique():
                safe_project = sanitize_filename(project)
                df_proj = df[df['Project name'] == project]
                
                if 'Workcentre' in df_proj.columns and not df_proj['Workcentre'].empty:
                    wc_summary = df_proj.groupby('Workcentre')['Hours'].sum().sort_values(ascending=False)
                    if not wc_summary.empty and wc_summary.sum() > 0:
                        fig, ax = plt.subplots(figsize=(10, 5))
                        bars = ax.barh(wc_summary.index, wc_summary.values, color='skyblue')
                        ax.bar_label(bars, labels=[f"{v:.1f}" for v in wc_summary.values], padding=3)
                        ax.set_title(f"{project} - Hours by Workcentre")
                        wc_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
                        plt.tight_layout()
                        fig.savefig(wc_path, dpi=150)
                        plt.close(fig)
                        charts_for_pdf.append((wc_path, f"{project} - Hours by Workcentre", project))
                        
                if 'Task' in df_proj.columns and not df_proj['Task'].empty:
                    task_summary = df_proj.groupby('Task')['Hours'].sum().sort_values(ascending=False)
                    if not task_summary.empty and task_summary.sum() > 0:
                        fig, ax = plt.subplots(figsize=(10, 6))
                        bars = ax.barh(task_summary.index, task_summary.values, color='lightgreen')
                        ax.bar_label(bars, labels=[f"{v:.1f}" for v in task_summary.values], padding=3)
                        ax.set_title(f"{project} - Hours by Task")
                        task_path = os.path.join(tmp_dir, f"{safe_project}_task.png")
                        plt.tight_layout()
                        fig.savefig(task_path, dpi=150)
                        plt.close(fig)
                        charts_for_pdf.append((task_path, f"{project} - Hours by Task", project))

                df_proj_weekend = df_proj[df_proj['IsWeekend'] == True]
                if not df_proj_weekend.empty:
                    ot_summary = df_proj_weekend.groupby('Task')['Hours'].sum().sort_values(ascending=False)
                    if not ot_summary.empty and ot_summary.sum() > 0:
                        fig, ax = plt.subplots(figsize=(10, 5))
                        bars = ax.barh(ot_summary.index, ot_summary.values, color='salmon')
                        ax.bar_label(bars, labels=[f"{v:.1f}" for v in ot_summary.values], padding=3)
                        ax.set_title(f"{project} - Hours by Task (Weekend OT)")
                        ot_path = os.path.join(tmp_dir, f"{safe_project}_weekend_ot.png")
                        plt.tight_layout()
                        fig.savefig(ot_path, dpi=150)
                        plt.close(fig)
                        charts_for_pdf.append((ot_path, f"{project} - Weekend Overtime by Task", project))

        pdf_config_info = {
            "Mode": config.get('mode', 'N/A').capitalize(), "Year": str(config.get('year', '')), "Months": ', '.join(config.get('months', [])) if config.get('months') else "Tất cả",
            "Project": ', '.join(config['project_filter_df'][config['project_filter_df']['Include'] == 'yes']['Project Name'].tolist()) if 'project_filter_df' in config and not config['project_filter_df'].empty else "Không có"
        }
        create_pdf_from_charts_comp(charts_for_pdf, pdf_report_path, "TRIAC TIME REPORT - STANDARD", pdf_config_info, logo_path)
        return True
    except Exception as e:
        traceback.print_exc()
        return False
    finally:
        if os.path.exists(tmp_dir): shutil.rmtree(tmp_dir)

def create_pdf_from_charts_comp(charts_data, output_path, title, config_info, logo_path_inner, filter_mode="Total"):
    from collections import defaultdict
    from PIL import Image
    today_str = datetime.today().strftime('%Y-%m-%d')
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_font('DejaVu', '', 'font/dejavu-fonts-ttf-2.37/ttf/DejaVuSans.ttf', uni=True)
    pdf.add_font('DejaVu', 'B', 'font/dejavu-fonts-ttf-2.37/ttf/DejaVuSans-Bold.ttf', uni=True)

    pdf.set_font('DejaVu', 'B', 16)
    pdf.add_page()
    if os.path.exists(logo_path_inner): pdf.image(logo_path_inner, x=10, y=10, w=30)
    pdf.ln(35)
    pdf.cell(0, 10, title, ln=True, align='C')
    pdf.set_font("DejaVu", '', 11)
    pdf.ln(5)
    pdf.cell(0, 10, f"Generated on: {today_str}", ln=True, align='C')
    pdf.ln(10)

    label_width, value_width, line_height = 40, 150, 8
    pdf.set_fill_color(240, 240, 240)
    for key, value in config_info.items():
        value_str = "N/A" if pd.isna(value) else str(value)
        value_lines = pdf.multi_cell(value_width, line_height, value_str, border=0, split_only=True)
        row_height = line_height * len(value_lines)
        x, y = pdf.get_x(), pdf.get_y()
        pdf.set_font("DejaVu", 'B', 11)
        pdf.multi_cell(label_width, row_height, key, border=1, fill=True)
        pdf.set_xy(x + label_width, y)
        pdf.set_font("DejaVu", '', 11)
        pdf.multi_cell(value_width, line_height, value_str, border=1)
        pdf.set_x(10)

    filter_mode_display = "By Task" if filter_mode == "Task" else ("By Workcentre" if filter_mode == "Workcentre" else "By Total Hours")
    pdf.ln(5)
    pdf.set_font("DejaVu", 'B', 11)
    pdf.cell(0, 8, f"Filter mode: {filter_mode_display}", ln=True)

    project_charts = defaultdict(list)
    for img_path, chart_title, project_name in charts_data:
        project_charts[project_name].append((img_path, chart_title))

    for project_name, charts in project_charts.items():
        for img_path, chart_title in charts:
            if not os.path.exists(img_path): continue
            img = Image.open(img_path)
            img_width, img_height = img.size
            aspect_ratio = img_height / img_width

            margin = 10
            pdf.add_page(orientation='L' if img_width > img_height else 'P')
            page_w, page_h = (297, 210) if img_width > img_height else (210, 297)

            pdf.set_font("DejaVu", 'B', 12)
            if os.path.exists(logo_path_inner): pdf.image(logo_path_inner, x=10, y=8, w=25)
            pdf.set_y(35)
            pdf.cell(0, 6, f"Project: {project_name}" if project_name else "Summary Charts", ln=True, align='C')
            pdf.set_font("DejaVu", '', 11)
            pdf.ln(0.5)
            pdf.cell(0, 2, chart_title, ln=True, align='C')

            max_w = page_w - 2 * margin
            new_w = max_w
            new_h = new_w * aspect_ratio
            if new_h > (page_h - 2 * margin):
                new_h = page_h - 2 * margin
                new_w = new_h / aspect_ratio
            pdf.image(img_path, x=(page_w - new_w) / 2, y=pdf.get_y() + 1.5, w=new_w, h=new_h)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    pdf.output(output_path, "F")
    return True, "✅ PDF created"

def create_comparison_chart(df, mode, title, x_label, y_label, path, config, filter_mode="Total"):
    output_dir = "tmp_comparison"
    try:
        os.makedirs(output_dir, exist_ok=True)
        charts = {}
        df = df.copy()

        if filter_mode == "Task": df = df[df['Task'] != 'All']
        elif filter_mode == "Workcentre": df = df[df['Workcentre'] != 'All']
        elif filter_mode == "Total":
            df.loc[:, 'Task'] = 'All'
            df.loc[:, 'Workcentre'] = 'All'

        if df.empty: return {}
        if 'MonthName' in df.columns:
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            df['MonthName'] = pd.Categorical(df['MonthName'], categories=month_order, ordered=True)

        if 'Year' in df.columns and 'MonthName' in df.columns:
            df['YearMonth'] = df['Year'].astype(str) + "-" + df['MonthName'].astype(str)
            df_sorted = df.groupby(['Project Name', 'Year', 'MonthName', 'YearMonth'], as_index=False)['Total Hours'].sum()

            projects = df_sorted['Project Name'].unique()
            all_yearmonths = sorted(df_sorted['YearMonth'].unique())
            x = np.arange(len(all_yearmonths))
            width = 0.8 / len(projects) if len(projects) > 1 else 0.6

            fig, ax = plt.subplots(figsize=(15, 8.3))
            for i, project in enumerate(projects):
                df_proj = df_sorted[df_sorted['Project Name'] == project]
                y_vals = []
                for ym in all_yearmonths:
                    match = df_proj[df_proj['YearMonth'] == ym]
                    y_vals.append(match['Total Hours'].sum() if not match.empty else 0)
                ax.bar(x + i * width, y_vals, width=width, label=project)
                for j, val in enumerate(y_vals):
                    if val > 0: ax.annotate(f"{val:.0f}", xy=(x[j] + i * width, val), xytext=(0, 5), textcoords="offset points", ha='center', fontsize=8, rotation=90)

            ax.set_title(f"{title} - Over Time")
            ax.set_xlabel(x_label)
            ax.set_ylabel(y_label)
            ax.set_xticks(x + width * (len(projects) - 1) / 2)
            ax.set_xticklabels(all_yearmonths, rotation=45, ha='right')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.20), ncol=5, fontsize=8)
            plt.tight_layout()
            chart_path = os.path.join(output_dir, "chart_time.png")
            fig.savefig(chart_path, dpi=150)
            plt.close(fig)
            charts["time"] = chart_path

        if 'Task' in df.columns and filter_mode == "Task":
            df_task = df.groupby(['Task', 'Project Name'], as_index=False)['Total Hours'].sum()
            if not df_task.empty:
                df_pivot = df_task.pivot(index='Task', columns='Project Name', values='Total Hours').fillna(0)
                fig, ax = plt.subplots(figsize=(11.7, 8.3))
                bars = df_pivot.plot(kind='bar', ax=ax)
                for container in bars.containers:
                    for bar in container:
                        height = bar.get_height()
                        if height > 0: ax.annotate(f"{height:.0f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 5), textcoords="offset points", ha='center', fontsize=8, rotation=90)
                ax.set_title(f"{title} - By Task")
                ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
                ax.legend(title="Project Name", loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=4, fontsize=8, frameon=False)
                plt.tight_layout()
                chart_path = os.path.join(output_dir, "chart_task.png")
                fig.savefig(chart_path, dpi=150)
                plt.close(fig)
                charts["task"] = chart_path
                
        if 'Workcentre' in df.columns and filter_mode == "Workcentre":
            df_wc = df.groupby(['Workcentre', 'Project Name'], as_index=False)['Total Hours'].sum()
            if not df_wc.empty:
                df_pivot = df_wc.pivot(index='Workcentre', columns='Project Name', values='Total Hours').fillna(0)
                fig, ax = plt.subplots(figsize=(15, 8.3))
                df_pivot.plot(kind='bar', ax=ax)
                ax.set_title(f"{title} - By Workcentre")
                ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
                handles, labels = ax.get_legend_handles_labels()
                if ax.get_legend(): ax.get_legend().remove()
                fig.legend(handles, labels, loc='lower center', bbox_to_anchor=(0.5, -0.15), ncol=min(len(labels), 5), fontsize=8, frameon=False)
                fig.subplots_adjust(left=0.08, right=0.98, top=0.75, bottom=0.33)
                chart_path = os.path.join(output_dir, "chart_workcentre.png")
                fig.savefig(chart_path, dpi=150, bbox_inches='tight')
                plt.close(fig)
                charts["workcentre"] = chart_path

        if filter_mode == "Total":
            df_total = df.groupby("Project Name", as_index=False)["Total Hours"].sum()
            if not df_total.empty:
                fig, ax = plt.subplots(figsize=(15.7, 8.3))
                bars = ax.bar(df_total["Project Name"], df_total["Total Hours"])
                ax.set_title(f"{title} - Total Hours by Project")
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                chart_path = os.path.join(output_dir, "chart_total.png")
                fig.savefig(chart_path, dpi=150)
                plt.close(fig)
                charts["total"] = chart_path

        return charts
    except Exception as e:
        print(f"Chart error: {e}")
        return None

def export_comparison_pdf_report(df_comparison, comparison_config, pdf_file_path, comparison_mode, logo_path, filter_mode="Total"):
    if 'Hours' not in df_comparison.columns: return False, "Thiếu cột Hours"
    if df_comparison.empty: return False, "Dữ liệu rỗng"
    tmp_dir = tempfile.mkdtemp()
    try:
        success, msg = generate_comparison_pdf_report(df_comparison, comparison_config, pdf_file_path, comparison_mode, logo_path, filter_mode)
        return success, msg
    except Exception as e:
        return False, f"❌ Lỗi PDF: {e}"
    finally:
        if os.path.exists(tmp_dir): shutil.rmtree(tmp_dir)

def generate_comparison_pdf_report(df_comparison, comparison_config, pdf_file_path, comparison_mode, logo_path, filter_mode="Total"):
    tmp_dir = "tmp_comparison"
    os.makedirs(tmp_dir, exist_ok=True)
    charts_for_pdf = []

    try:
        filtered_projects = comparison_config.get("filtered_projects", [])
        pdf_config_info = {"Mode": comparison_mode, "Year": ', '.join(map(str, comparison_config.get('years', []))) or "N/A", "Months": ', '.join(comparison_config.get('months', [])) or "All", "Projects": ', '.join(filtered_projects) or "Không có"}
        chart_title = f"So sánh dự án"
        chart_path_placeholder = os.path.join(tmp_dir, "unused.png")
        charts_dict = create_comparison_chart(df_comparison, comparison_mode, chart_title, "Dự án", "Giờ", chart_path_placeholder, comparison_config, filter_mode)
        
        if charts_dict:
            chart_title_map = {"time": "So sánh giờ theo thời gian", "total": "Tổng giờ theo từng dự án", "task": "So sánh giờ theo Task giữa các dự án", "workcentre": "So sánh giờ theo Workcentre giữa các dự án"}
            for key in ["time", "total", "task", "workcentre"]:
                chart_path = charts_dict.get(key)
                if chart_path and os.path.exists(chart_path):
                    charts_for_pdf.append((chart_path, chart_title_map.get(key, key), "Tổng hợp nhiều dự án"))
        else:
            return False, "⚠️ Không tạo được biểu đồ"
            
        success, msg = create_pdf_from_charts_comp(charts_for_pdf, pdf_file_path, "TRIAC TIME REPORT - COMPARISON", pdf_config_info, logo_path, filter_mode=filter_mode)
        return success, msg
    except Exception as e:
        return False, f"❌ Exception: {e}"
    finally:
        if os.path.exists(tmp_dir): shutil.rmtree(tmp_dir)

def apply_comparison_filters(df_raw, comparison_config, comparison_mode, filter_mode="Total"):
    if not isinstance(df_raw, pd.DataFrame): return pd.DataFrame(), "Dữ liệu lỗi.", []   
    years = list(comparison_config.get('years', []))
    months = list(comparison_config.get('months', []))
    selected_projects = [p for p in comparison_config.get('selected_projects', []) if str(p).strip()]

    df_filtered = df_raw.copy()
    df_filtered['Hours'] = pd.to_numeric(df_filtered['Hours'], errors='coerce').fillna(0)

    if years: df_filtered = df_filtered[df_filtered['Year'].isin(years)]
    if months: df_filtered = df_filtered[df_filtered['MonthName'].isin(months)]
        
    df_filtered_projects = df_filtered['Project name'].unique().tolist()
    selected_projects = [p for p in selected_projects if p in df_filtered_projects]
    comparison_config["filtered_projects"] = selected_projects
    
    if selected_projects: df_filtered = df_filtered[df_filtered['Project name'].isin(selected_projects)]
    else: return pd.DataFrame(), "Vui lòng chọn dự án.", []
    if df_filtered.empty: return pd.DataFrame(), f"Dữ liệu trống.", []

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        df_comparison = df_filtered.copy()
        df_comparison.rename(columns={'Project name': 'Project Name'}, inplace=True)
        df_comparison['Total Hours'] = df_comparison['Hours']
        df_comparison['Task'] = df_comparison.get('Task', 'All')
        df_comparison['Workcentre'] = df_comparison.get('Workcentre', 'All')
        return df_comparison, f"So sánh dự án trong tháng", selected_projects

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        df_pivot = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        df_comparison = df_pivot.reset_index()
        df_comparison['Total Hours'] = df_comparison.sum(numeric_only=True, axis=1)
        df_comparison.rename(columns={'Project name': 'Project Name'}, inplace=True)
        return df_comparison, f"So sánh dự án trong năm", selected_projects

    elif comparison_mode in ["So Sánh Nhiều Dự Án Qua Các Tháng/Năm", "Compare Projects Over Time (Months/Years)"]:
        df_comparison = df_filtered.copy()
        df_comparison.rename(columns={'Project name': 'Project Name'}, inplace=True)
        df_comparison['Total Hours'] = df_comparison['Hours']
        return df_comparison, "So sánh nhiều dự án qua các năm và tháng", selected_projects

    return pd.DataFrame(), "❌ Không hỗ trợ.", []

def export_comparison_report(df_comparison, comparison_config, output_file_path, comparison_mode, filter_mode="Total"):
    try:
        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            if df_comparison.empty:
                pd.DataFrame({"Message": ["Không có dữ liệu"]}).to_excel(writer, sheet_name='Comparison Report', index=False)
            else:
                df_comparison.to_excel(writer, sheet_name='Comparison Report', index=False)  

            wb = writer.book
            ws = wb['Comparison Report']
            info_row = ws.max_row + 2 
            ws.cell(row=info_row, column=1, value=f"BÁO CÁO SO SÁNH: {comparison_mode}").font = ws.cell(row=info_row, column=1).font.copy(bold=True)

            if not df_comparison.empty and 'Total Hours' in df_comparison.columns:
                chart = BarChart()
                chart.title = "Báo cáo so sánh"
                df_chart_data = df_comparison[df_comparison['Project Name'] != 'Total'] if 'Project Name' in df_comparison.columns else df_comparison
                max_r = 2 + len(df_chart_data) - 1
                col_idx = df_comparison.columns.get_loc('Total Hours') + 1
                chart.add_data(Reference(ws, min_col=col_idx, min_row=2, max_row=max_r), titles_from_data=False)
                chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=max_r))
                ws.add_chart(chart, f"A{info_row + 5}")

        wb.save(output_file_path)
        return True
    except Exception as e:
        print(f"Lỗi xuất báo cáo so sánh Excel: {e}")
        return False
