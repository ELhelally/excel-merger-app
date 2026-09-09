import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import pandas as pd
import os
import glob
import zipfile
import datetime
import io

# إعدادات صفحة الموقع
st.set_page_config(page_title="دمج كشوف الحضور والانصراف", page_icon="📊", layout="wide")

# عنوان الموقع
st.title("📊 أداة دمج وتنسيق كشوف الحضور والانصراف")
st.write("قم برفع ملف الـ ZIP الذي يحتوي على شيتات الموظفين، حدد اسم الملف الناتج، ثم اضغط على زر المعالجة للحصول على ملفك المجمع والمنسق.")

# 1. رفع الملف
uploaded_zip = st.file_uploader("اختر ملف الـ ZIP (مثل: employees.zip)", type=["zip"])

# 2. تحديد اسم الملف الناتج
output_custom_name = st.text_input("📝 اكتب اسم ملف الإكسل الناتج (بدون إضافة .xlsx):", value="كشف_حضور_وانصراف_شهر_مارس_المجمع")

def process_employee_sheet(file_path):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        all_rows = list(sheet.iter_rows(values_only=True))
        
        emp_name, emp_id, emp_role = None, None, None
        header_row_idx = None
        
        for r_idx, row in enumerate(all_rows):
            if any('اليوم' in str(c) for c in row if c) and any('التاريخ' in str(c) for c in row if c):
                header_row_idx = r_idx
                break
                
            for cell in row:
                if cell is None: continue
                cell_str = str(cell).strip()
                
                if 'الاسم' in cell_str and not emp_name:
                    emp_name = cell_str.split(':')[-1].strip() if ':' in cell_str else None
                    if not emp_name:
                        idx = list(row).index(cell)
                        if idx + 1 < len(row): emp_name = str(row[idx+1]).strip()

                if any(k in cell_str for k in ['رقم العامل', 'رقم الموظف', 'الكود']) and not emp_id:
                    emp_id = cell_str.split(':')[-1].strip() if ':' in cell_str else None
                    if not emp_id:
                        idx = list(row).index(cell)
                        if idx + 1 < len(row): emp_id = str(row[idx+1]).strip()

                if 'الوظيفة' in cell_str and not emp_role:
                    emp_role = cell_str.split(':')[-1].strip() if ':' in cell_str else None
                    if not emp_role:
                        idx = list(row).index(cell)
                        if idx + 1 < len(row): emp_role = str(row[idx+1]).strip()

        fname = os.path.basename(file_path)
        if '-' in fname:
            f_parts = fname.replace('.xlsx', '').replace('.xls', '').split('-')
            if not emp_name and len(f_parts) >= 1: emp_name = f_parts[0].strip()
            if not emp_id and len(f_parts) >= 2: emp_id = f_parts[1].strip()

        rows_data = []
        if header_row_idx is not None:
            for r in all_rows[header_row_idx + 1:]:
                if not any(r): continue
                
                day_str = str(r[0]).strip() if r[0] is not None else ""
                
                if any(term in day_str for term in ['إجمالي', 'اجمالي', 'ملخص', 'عدد أيام', 'التقدير العام']):
                    if 'إجمالي' in day_str or 'اجمالي' in day_str: break
                    continue
                    
                if r[0] is None and r[1] is None: continue

                date_val = r[1]
                if isinstance(date_val, (datetime.datetime, datetime.date)):
                    date_clean = date_val.strftime("%Y-%m-%d")
                elif date_val is not None:
                    date_clean = str(date_val).split(' ')[0].strip()
                else:
                    date_clean = None

                rows_data.append({
                    'اليوم': r[0] if len(r) > 0 else None,
                    'التاريخ': date_clean,
                    'رقم العامل': emp_id,
                    'الاسم': emp_name,
                    'الوظيفة': emp_role,
                    'رقم العملية': r[3] if len(r) > 3 else None,
                    'العملية': r[4] if len(r) > 4 else None,
                    'المكان': r[5] if len(r) > 5 else None,
                    'من': r[6] if len(r) > 6 else None,
                    'إلى': r[7] if len(r) > 7 else None,
                    'الإنتقالات': r[8] if len(r) > 8 else None,
                    'بدل غذاء': r[9] if len(r) > 9 else None,
                    'توقيع المشرف': r[10] if len(r) > 10 else None,
                    'ملاحظات': r[11] if len(r) > 11 else None
                })
        return pd.DataFrame(rows_data)
    except Exception as e:
        st.error(f"خطأ في الملف {file_path}: {e}")
        return pd.DataFrame()

if uploaded_zip is not None:
    if st.button("🚀 بدء دمج الملفات وتنسيقها"):
        with st.spinner("جاري فك الضغط ومعالجة الشيتات..."):
            extract_dir = "./temp_extracted"
            os.makedirs(extract_dir, exist_ok=True)
            
            with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                
            all_files = glob.glob(f"{extract_dir}/**/*.xlsx", recursive=True) + glob.glob(f"{extract_dir}/**/*.xls", recursive=True)
            
            dfs = [process_employee_sheet(f) for f in all_files]
            master_df = pd.concat(dfs, ignore_index=True)
            
            cols_order = ['اليوم', 'التاريخ', 'رقم العامل', 'الاسم', 'الوظيفة', 'رقم العملية', 'العملية', 'المكان', 'من', 'إلى', 'الإنتقالات', 'بدل غذاء', 'توقيع المشرف', 'ملاحظات']
            for col in cols_order:
                if col not in master_df.columns:
                    master_df[col] = None
            master_df = master_df[cols_order]
            
            output_path = "temp_output.xlsx"
            master_df.to_excel(output_path, index=False)
            
            wb = openpyxl.load_workbook(output_path)
            ws = wb.active
            ws.sheet_view.rightToLeft = True
            
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            holiday_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
            thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
            align_center = Alignment(horizontal='center', vertical='center')
            
            for col_num in range(1, len(cols_order) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center

            for row_num in range(2, ws.max_row + 1):
                day_val = str(ws.cell(row=row_num, column=1).value or '')
                process_val = str(ws.cell(row=row_num, column=7).value or '')
                is_holiday = ('الجمعة' in day_val or 'الجمعه' in day_val or 'إجازة' in process_val or 'اجازة' in process_val)
                
                for col_num in range(1, len(cols_order) + 1):
                    cell = ws.cell(row=row_num, column=col_num)
                    cell.alignment = align_center
                    cell.border = thin_border
                    if col_num == 2 and cell.value:
                        cell.number_format = 'yyyy-mm-dd'
                    if is_holiday:
                        cell.fill = holiday_fill

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            # تنظيف اسم الملف وإضافة الامتداد
            clean_filename = output_custom_name.strip()
            if not clean_filename.endswith(".xlsx"):
                clean_filename += ".xlsx"
            
            st.success(f"✅ تم دمج {len(all_files)} ملف بإجمالي {len(master_df)} صف بنجاح!")
            
            # زر التحميل بالاسم الذي حدده المستخدم
            st.download_button(
                label=f"📥 تحميل الملف: {clean_filename}",
                data=output_buffer,
                file_name=clean_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
