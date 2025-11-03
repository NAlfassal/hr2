*** Settings ***
Documentation     Send Form, Reminders, and Schedule Monthly Reports Automatically
Library           OperatingSystem
Library           DateTime
Library           RPA.Email.ImapSmtp
Library           RPA.Tables
Library           Collections

*** Variables ***
# --- إعدادات بايثون والملفات ---
${PYTHON_EXE}          D:${/}employee form${/}venv${/}Scripts${/}python.exe
${PROCESSOR_SCRIPT}    ${CURDIR}${/}report_processor.py  
${CLEANER_SCRIPT}      ${CURDIR}${/}excel_cleaner.py  
${REMINDER_SCRIPT}     ${CURDIR}${/}reminder_checker.py
${RESPONSES_FILE}      D:${/}employee form${/}data${/}responses.xlsx
${MASTER_LIST_FILE}    D:${/}employee form${/}data${/}Employees_emails_list.xlsx

# --- إعدادات الإيميل ---
${GMAIL_USER}     nadaalfassal@gmail.com
${GMAIL_PASS}     
${SMTP_SERVER}    smtp.gmail.com
${SMTP_PORT}      587
${FORM_URL}       http://localhost:5000/form
${HR_EMAIL}       nadaalfassal@gmail.com
@{EMPLOYEES}      nadaalfassal@gmail.com

# --- 🔥 تعديل: تحديد كل التواريخ والأوقات بشكل صريح ---
${FORM_SEND_DATE}       2025-11-03  # تاريخ إرسال الفورم الأولي
${REMINDER_SEND_DATE}   2025-11-03  # تاريخ إرسال التذكيرات
${FINAL_REPORT_DATE}    2025-11-03  # تاريخ إرسال التقرير النهائي

${FORM_TARGET_TIME}     15:38:00    # وقت إرسال الفورم الأولي
${REMINDER_TIME}        15:40:00    # وقت إرسال التذكيرات
${REPORT_TARGET_TIME}   15:42:00    # وقت إرسال التقرير النهائي

*** Tasks ***
Execute Full Automation Cycle
    Log To Console    Automation started.

    # 1. إرسال الفورم في التاريخ المحدد
    Wait For Specific Time    ${FORM_SEND_DATE} ${FORM_TARGET_TIME}
    Send Form Link To Employees

    # 2. 🔥 تعديل: إرسال التذكيرات في التاريخ المحدد (بدون حساب)
    Wait For Specific Time    ${REMINDER_SEND_DATE} ${REMINDER_TIME}
    Send Reminders To Non-Responders

    # 3. إرسال التقرير النهائي في التاريخ المحدد
    Wait For Specific Time    ${FINAL_REPORT_DATE} ${REPORT_TARGET_TIME}
    Copy Send And Reset Monthly Data 


*** Keywords ***
Wait For Specific Time
    [Arguments]    ${target_datetime_str}
    
    ${target_datetime_obj}=    Convert Date    ${target_datetime_str}    date_format=%Y-%m-%d %H:%M:%S
    ${target_epoch}=           Convert Date    ${target_datetime_obj}    result_format=epoch
    
    ${current_epoch}=          Get Current Date    result_format=epoch

    ${sleep_seconds}=    Evaluate    ${target_epoch} - ${current_epoch}
    
    ${sleep_seconds}=    Convert To Integer    ${sleep_seconds}

    IF    ${sleep_seconds} > 0

        ${sleep_duration}=    Convert Time    ${sleep_seconds}s
        Log To Console    Waiting for ${sleep_duration} until ${target_datetime_str}...
        Sleep    ${sleep_seconds}s
    ELSE
        Log To Console    Target time ${target_datetime_str} has already passed or is current. Proceeding immediately.
    END

Authorize Gmail
    Authorize    account=${GMAIL_USER}    password=${GMAIL_PASS}    smtp_server=${SMTP_SERVER}    smtp_port=${SMTP_PORT}
    Log To Console    Successfully logged into Gmail ✅

Send Form Link To Employees
    # Re-authorize connection before the first send
    Authorize Gmail 
    Log To Console    Sending form link to employees: ${FORM_URL}
    FOR    ${email}    IN    @{EMPLOYEES}
        Send Message
        ...    sender=${GMAIL_USER}
        ...    recipients=${email}
        ...    subject=📋 حصر الأنشطة المعرفية في الهيئة العامة للغذاء والدواء
        ...    body=مرحبًا، الرجاء تعبئة النموذج عبر الرابط التالي:\n${FORM_URL}
    END
    Log To Console    ✅ Link sent successfully.

Send Reminders To Non-Responders
    [Documentation]    Runs the Python script to check and send reminders.
    Log To Console    Starting reminder check for non-responders...

    # الأمر الكامل لتشغيل سكريبت التذكير مع كل المتغيرات المطلوبة
    ${REMINDER_COMMAND}=    Catenate    SEPARATOR=    "${PYTHON_EXE}" -X utf8 "${REMINDER_SCRIPT}" "${RESPONSES_FILE}" "${MASTER_LIST_FILE}" "${GMAIL_USER}" "${GMAIL_PASS}" "${FORM_URL}" 2>&1
    
    ${REMINDER_RESULT}=    OperatingSystem.Run    ${REMINDER_COMMAND}
    Log To Console    Reminder Script Output: ${REMINDER_RESULT}

    ${error_found}=    Run Keyword And Return Status    Should Contain    ${REMINDER_RESULT}    error    ignore_case=True
    IF    ${error_found}
        Log To Console    An error occurred in the reminder script, but continuing execution. Error: ${REMINDER_RESULT}    level=WARN
    END

Copy Send And Reset Monthly Data
    [Documentation]    Processes, copies, sends, and finally resets the original data file.
    
    # 1. تشغيل المعالج: ينظف ويظلل التكرارات في الملف الأصلي
    Log To Console    Starting Data Processing (Clean & Highlight Duplicates)...
    
    ${PROCESS_COMMAND}=    Catenate    SEPARATOR=    "${PYTHON_EXE}" "${PROCESSOR_SCRIPT}" "${RESPONSES_FILE}" 2>&1
    ${PROCESSOR_RESULT}=    OperatingSystem.Run    ${PROCESS_COMMAND}
    
    Log To Console    Processor Script Output: ${PROCESSOR_RESULT}
    
    ${error_found_in_processor}=    Run Keyword And Return Status    Should Contain    ${PROCESSOR_RESULT}    ERROR    ignore_case=True
    IF    ${error_found_in_processor}
        FAIL    Data Processor Failed: ${PROCESSOR_RESULT}
    END

    # 2. تحديد مسار النسخ المؤقتة
    ${CURRENT_DATE}=    Get Current Date    result_format=%Y-%m-%d_%H%M%S
    ${TEMP_FILENAME}=    Set Variable    temp_responses_${CURRENT_DATE}.xlsx
    ${TEMP_COPY_PATH}=    Set Variable    ${CURDIR}${/}${TEMP_FILENAME} 
    
    # 3. نسخ الملف الأصلي (الآن هو ملف مُعالج ومظلل)
    Log To Console    Creating temporary copy of the PROCESSED file...
    OperatingSystem.Copy File    ${RESPONSES_FILE}    ${TEMP_COPY_PATH}
    Log To Console    ✅ Temporary copy created at: ${TEMP_COPY_PATH}
    
    # 4. إرسال النسخة المؤقتة إلى الموارد البشرية
    Send Archive To HR    ${TEMP_COPY_PATH}
    
    # 5. حذف النسخة المؤقتة
    OperatingSystem.Remove File    ${TEMP_COPY_PATH}
    Log To Console    ✅ Temporary copy deleted.
    
    # 6. تشغيل سكريبت المسح النهائي
    Log To Console    Starting final file reset (clearing content)...
    
    ${CLEAN_COMMAND}=     Catenate    SEPARATOR=    "${PYTHON_EXE}" "${CLEANER_SCRIPT}" "${RESPONSES_FILE}" 2>&1
    ${CLEANER_RESULT}=    OperatingSystem.Run    ${CLEAN_COMMAND}
    
    Log To Console    Cleaner Script Output: ${CLEANER_RESULT}
    
    ${error_found_in_cleaner}=    Run Keyword And Return Status    Should Contain    ${CLEANER_RESULT}    ERROR    ignore_case=True
    IF    ${error_found_in_cleaner}
        FAIL    Excel Cleaner Failed: ${CLEANER_RESULT}
    END
    
    Log To Console    ✅ Original responses file reset successfully.

Send Archive To HR
    [Arguments]    ${attachment_path}
    # Re-authorize connection before the second send
    Authorize Gmail 
    Log To Console    Sending monthly archived data to HR...
    Send Message
    ...    sender=${GMAIL_USER}
    ...    recipients=${HR_EMAIL}
    ...    subject=📊 التقرير الشهري للأنشطة المعرفية 
    ...    body=مرفق ملف الردود المجمعة والمُعالجة، مع تظليل الصفوف المكررة باللون الأحمر للمراجعة.
    ...    attachments=${attachment_path}

    Log To Console    Report sent successfully to HR.
