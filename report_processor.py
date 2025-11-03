# -*- coding: utf-8 -*-
import pandas as pd
import sys
import os
import re

def normalize_text(text):
    if pd.isna(text):
        return text
    text = str(text).strip()
    text = re.sub('[أإآ]', 'ا', text)
    return text

def process_and_highlight_duplicates(file_path):
    try:
        if not os.path.exists(file_path):
            print(f"Error: File not found at {file_path}", file=sys.stderr)
            return

        print("Starting data cleaning and duplicate highlighting...")
        df = pd.read_excel(file_path)
        
        # 🔥 تعديل حاسم: تحديث قائمة الأعمدة المفتاحية
        # تم استبدال 'تاريخ النشاط' بالعمودين الجديدين
        # وتم تحديث بقية الأسماء لتطابق الفورم الجديد
        key_columns = [
            'تاريخ بداية النشاط', 
            'تاريخ نهاية النشاط',
            'اسم القطاع', 
            'الإدارة التنفيذية', 
            'الإدارة',
            'موضوع النشاط', 
            'اسم المقدم', 
            'الفئة المستهدفة (تفاصيل)'
        ]
        
        # التأكد من وجود الأعمدة قبل استخدامها
        existing_key_columns = [col for col in key_columns if col in df.columns]
        if not existing_key_columns:
            print("Warning: No key columns found for duplicate check. Skipping.")
            return

        # تطبيق التنظيف على أعمدة النصوص
        text_cols_to_clean = ['موضوع النشاط', 'اسم المقدم', 'الفئة المستهدفة (تفاصيل)']
        for col in text_cols_to_clean:
            if col in df.columns:
                df[col] = df[col].apply(normalize_text)

        # تحديد التكرارات بناءً على الأعمدة الموجودة
        df['is_duplicate'] = df.duplicated(subset=existing_key_columns, keep='first')
        
        if not df['is_duplicate'].any():
            print("No duplicates found.")
            # لا داعي لإعادة الحفظ إذا لم يتغير شيء
            return

        print(f"Found {df['is_duplicate'].sum()} duplicate rows. Highlighting in red...")

        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        df.drop(columns=['is_duplicate']).to_excel(writer, sheet_name='الردود', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['الردود']
        worksheet.right_to_left()

        red_format = workbook.add_format({'bg_color': '#FFC7CE'})
        
        for row_index in df.index:
            if df.loc[row_index, 'is_duplicate']:
                excel_row = row_index + 1
                worksheet.set_row(excel_row, None, red_format)
                
        writer.close()
        
        print("Data cleaning and highlighting complete.")

    except Exception as e:
        print(f"Error during Excel processing: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Error: Excel file path not provided.", file=sys.stderr)
        sys.exit(1)
        
    excel_path = sys.argv[1]
    process_and_highlight_duplicates(excel_path)