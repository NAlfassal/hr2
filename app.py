# -*- coding: utf-8 -*-
from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime
import mimetypes # 🔥 إضافة جديدة: استيراد المكتبة

# 🔥 إضافة جديدة: هذا السطر يخبر بايثون بالنوع الصحيح لملفات .js
mimetypes.add_type('application/javascript', '.js')

# 🔥 هذا هو السطر الحاسم: يخبر فلاسك بمكان الملفات
app = Flask(__name__, static_folder='static', template_folder='.')

DATA_PATH = os.path.join("data", "responses.xlsx")

@app.route("/form", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        # (منطق حفظ البيانات يبقى كما هو)
        column_order = [
            "البريد الإلكتروني", "اسم القطاع", "الإدارة التنفيذية", "الإدارة", "القسم",
            "هل توجد أنشطة؟", "موضوع النشاط", "نوع النشاط", "هدف استراتيجي 1", 
            "هدف استراتيجي 2", "تصنيف المقدم", "تاريخ بداية النشاط", "تاريخ نهاية النشاط",
            "اسم المقدم", "مسؤول الحضور", "الفئة المستهدفة (النوع)", "الفئة المستهدفة (تفاصيل)",
            "عدد الحضور", "مدة النشاط (ساعة)", "مكان الحفظ", "تاريخ الإرسال"
        ]
        form_data = request.form
        has_activities = form_data.get('hasActivities') == 'yes'
        data = {
            "البريد الإلكتروني": [form_data.get("employeeEmail")],
            "اسم القطاع": [form_data.get('sector')],
            "الإدارة التنفيذية": [form_data.get('department')],
            "الإدارة": [form_data.get('division')],
            "القسم": [form_data.get('section')],
            "هل توجد أنشطة؟": ["نعم" if has_activities else "لا يوجد"],
            "تاريخ الإرسال": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            "موضوع النشاط": [form_data.get("activityTopic")] if has_activities else ["لا يوجد"],
            "نوع النشاط": [form_data.get("activityType")] if has_activities else ["لا يوجد"],
            "هدف استراتيجي 1": [form_data.get("strategicGoalLevel1")] if has_activities else ["لا يوجد"],
            "هدف استراتيجي 2": [form_data.get("strategicGoalLevel2")] if has_activities else ["لا يوجد"],
            "تصنيف المقدم": [form_data.get("presenterCategory")] if has_activities else ["لا يوجد"],
            "تاريخ بداية النشاط": [form_data.get("activityStartDate")] if has_activities else ["لا يوجد"],
            "تاريخ نهاية النشاط": [form_data.get("activityEndDate")] if has_activities else ["لا يوجد"],
            "اسم المقدم": [form_data.get("presenterName")] if has_activities else ["لا يوجد"],
            "مسؤول الحضور": [form_data.get("attendanceResponsible")] if has_activities else ["لا يوجد"],
            "الفئة المستهدفة (النوع)": [form_data.get("targetAudienceType")] if has_activities else ["لا يوجد"],
            "الفئة المستهدفة (تفاصيل)": [form_data.get("targetAudienceDetails")] if has_activities else ["لا يوجد"],
            "عدد الحضور": [form_data.get("attendeeCount")] if has_activities else [0],
            "مدة النشاط (ساعة)": [form_data.get("activityDuration")] if has_activities else [0],
            "مكان الحفظ": [form_data.get("contentLocation")] if has_activities else ["لا يوجد"],
        }
        df = pd.DataFrame(data)
        df = df.reindex(columns=column_order)
        if os.path.exists(DATA_PATH):
            try:
                old_df = pd.read_excel(DATA_PATH)
                df = pd.concat([old_df, df], ignore_index=True)
            except Exception as e:
                print(f"Error reading existing Excel file: {e}.")
        try:
            writer = pd.ExcelWriter(DATA_PATH, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='الردود', index=False)
            worksheet = writer.sheets['الردود']
            worksheet.right_to_left()
            writer.close() 
            return "✅ تم استلام النموذج وحفظه بنجاح!"
        except Exception as e:
            print(f"Error writing to Excel: {e}")
            return "❌ فشل في حفظ البيانات.", 500
    
    return render_template("index.html")

if __name__ == "__main__":
    os.makedirs("data", exist_ok=True)
    app.run(debug=True)