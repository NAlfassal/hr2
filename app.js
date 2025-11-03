import {
  buildStoreFromJson,
  listByType,
  listChildren,
  validate
} from './org.loader.js';

let store;

// 🔥 خطوة 1: تم نقل الدوال المساعدة إلى هنا (النطاق العام)
// الآن أصبحت متاحة لجميع أجزاء الكود
function resetSelect(sel, placeholder, disable = true) {
  sel.innerHTML = '';
  const opt = document.createElement('option');
  opt.value = '';
  opt.textContent = placeholder;
  sel.appendChild(opt);
  sel.disabled = disable;
}

function fillSelect(sel, items, placeholder) {
  resetSelect(sel, placeholder, items.length === 0);
  for (const it of items) {
    const o = document.createElement('option');
    o.value = it.id;
    o.textContent = it.name;
    sel.appendChild(o);
  }
  sel.disabled = items.length === 0;
}

function isNoItemList(items) {
  return items.length === 1 && items[0].name.trim() === 'لا يوجد';
}

// 🔥 خطوة 2: تعريف الدالة الرئيسية مرة واحدة فقط
async function initOrgLists() {
  const data = await fetch('./static/org.data.json').then(r => r.json());
  store = buildStoreFromJson(data);

  const errs = validate(store);
  if (errs.length) console.warn('STRUCTURE ERRORS:', errs);

  // الآن يمكن لهذه الدالة العثور على fillSelect و resetSelect
  const sectorEl = document.getElementById('sector');
  const departmentEl = document.getElementById('department');
  const divisionEl = document.getElementById('division');
  const sectionEl  = document.getElementById('section');

  fillSelect(sectorEl, listByType(store, 'lv1'), 'اختر القطاع');
  resetSelect(departmentEl, 'اختر الإدارة التنفيذية');
  resetSelect(divisionEl,   'اختر الإدارة');
  resetSelect(sectionEl,    'اختر القسم');
}

// 🔥 خطوة 3: هذا الكود ينتظر تحميل الصفحة ثم يبدأ كل شيء
document.addEventListener('DOMContentLoaded', () => {
  // ===== عناصر DOM =====
  const sectorEl = document.getElementById('sector');
  const departmentEl = document.getElementById('department');
  const divisionEl = document.getElementById('division');
  const sectionEl  = document.getElementById('section');
  const activityQuestionEl = document.getElementById('activityQuestion');
  const formFieldsEl = document.getElementById('formFields');
  const submitBtn = document.getElementById('submitBtn');
  const noActivitiesMessage = document.getElementById('noActivitiesMessage');
  const successMessage = document.getElementById('successMessage');
  const loadingEl = document.getElementById('loading');
  const knowledgeForm = document.getElementById('knowledgeForm');

  // ===== Required on/off for dynamic fields =====
  function setDynamicFieldsRequired(isRequired) {
    document.querySelectorAll('#formFields [data-required]').forEach((el) => {
      if (isRequired) el.setAttribute('required', '');
      else el.removeAttribute('required');
    });
  }

  // ===== تحديد شهر الحصر الحالي =====
  function updateCurrentMonth() {
    const now = new Date();
    const months = [
      'يناير','فبراير','مارس','أبريل','مايو','يونيو',
      'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'
    ];
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const lastMonthName = months[lastMonth.getMonth()];
    const lastMonthYear = lastMonth.getFullYear();
    document.getElementById('currentMonth').textContent =
      `شهر الحصر: ${lastMonthName} ${lastMonthYear}`;
  }
  updateCurrentMonth();

  // ===== ربط الأحداث بالقوائم المنسدلة =====
  sectorEl.addEventListener('change', () => {
    const sectorId = sectorEl.value;
    activityQuestionEl.style.display = 'none';
    formFieldsEl.style.display = 'none';
    submitBtn.style.display = 'none';
    noActivitiesMessage.style.display = 'none';
    resetSelect(departmentEl, 'اختر الإدارة التنفيذية');
    resetSelect(divisionEl,   'اختر الإدارة');
    resetSelect(sectionEl,    'اختر القسم');
    if (!sectorId) return;
    const lv2 = listChildren(store, sectorId, 'lv2');
    if (isNoItemList(lv2)) {
      fillSelect(departmentEl, lv2, 'لا يوجد');
      departmentEl.disabled = true;
      divisionEl.disabled = true;
      sectionEl.disabled = true;
      return;
    }
    fillSelect(departmentEl, lv2, 'اختر الإدارة التنفيذية');
  });

  departmentEl.addEventListener('change', function () {
    const deptId = this.value;
    resetSelect(divisionEl, 'اختر الإدارة');
    resetSelect(sectionEl,  'اختر القسم');
    activityQuestionEl.style.display = deptId ? 'block' : 'none';
    if (!deptId) return;
    const lv3 = listChildren(store, deptId, 'lv3');
    if (isNoItemList(lv3)) {
      fillSelect(divisionEl, lv3, 'لا يوجد');
      divisionEl.disabled = true;
      sectionEl.disabled = true;
      return;
    }
    fillSelect(divisionEl, lv3, 'اختر الإدارة');
    activityQuestionEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
  });

  divisionEl.addEventListener('change', function () {
    const divId = this.value;
    resetSelect(sectionEl, 'اختر القسم');
    if (!divId) return;
    const lv4 = listChildren(store, divId, 'lv4');
    if (isNoItemList(lv4)) {
      fillSelect(sectionEl, lv4, 'لا يوجد');
      sectionEl.disabled = true;
      return;
    }
    fillSelect(sectionEl, lv4, 'اختر القسم');
  });

  // ===== نعم / لا (زر الإرسال يعمل في الحالتين) =====
  document.querySelectorAll('input[name="hasActivities"]').forEach((radio) => {
    radio.addEventListener('change', function () {
      if (this.value === 'yes') {
        showFormFields();
        setDynamicFieldsRequired(true);
        noActivitiesMessage.style.display = 'none';
        submitBtn.style.display = 'block';
      } else {
        setDynamicFieldsRequired(false);
        formFieldsEl.innerHTML = '';
        formFieldsEl.style.display = 'none';
        noActivitiesMessage.style.display = 'block';
        submitBtn.style.display = 'block';
        noActivitiesMessage.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      document.querySelectorAll('input[name="hasActivities"]').forEach((r) => {
        r.parentElement.style.borderColor = 'transparent';
        r.parentElement.style.transform = 'scale(1)';
      });
      this.parentElement.style.borderColor = this.value === 'yes' ? '#4caf50' : '#ff9800';
      this.parentElement.style.transform = 'scale(1.05)';
    });
  });

  // ===== حقول النموذج الديناميكية =====
  function showFormFields() {
    formFieldsEl.innerHTML = `
      <div class="form-group">
        <label class="form-label">موضوع النشاط <span class="required">*</span></label>
        <textarea class="form-control" id="activityTopic" name="activityTopic" required data-required placeholder="كتابة عنوان النشاط المعرفي الذي تم تنفيذه، وشرح مبسط لمحتوى النشاط إن أمكن"></textarea>
      </div>
      <div class="form-group">
        <label class="form-label">نوع النشاط <span class="required">*</span></label>
        <select class="form-control" id="activityType" name="activityType" required data-required>
          <option value="">اختر نوع النشاط</option>
          <option value="training-course">دورة تدريبية</option>
          <option value="workshop">ورشة عمل</option>
          <option value="lecture">محاضرة</option>
          <option value="seminar">ندوة</option>
          <option value="knowledge-meetings">اجتماعات معرفية</option>
          <option value="scientific-meeting">لقاء علمي</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">الهدف الإستراتيجي من المستوى الأول <span class="required">*</span></label>
        <select class="form-control" id="strategicGoalLevel1" name="strategicGoalLevel1" required data-required>
          <option value="">اختر الهدف الإستراتيجي من المستوى الأول</option>
          <option value="develop-regulatory-system">تطوير المنظومة الرقابية</option>
          <option value="improve-communication-awareness">تحسين التواصل والتوعية</option>
          <option value="enhance-product-availability">تعزيز توفر المنتجات</option>
          <option value="enhance-international-leadership">تعزيز الريادة الدولية</option>
          <option value="diversify-revenue-sources">تنويع مصادر الإيرادات</option>
          <option value="develop-biotech-regulation">تطوير التشريع والرقابة على منتجات التقنية الحيوية والحديثة</option>
          <option value="support-research-innovation">دعم البحث والابتكار</option>
          <option value="enable-investor">تمكين المستثمر</option>
          <option value="develop-human-capital">تنمية رأس المال البشري</option>
          <option value="increase-digital-tech-usage">زيادة استخدام التقنيات الرقمية المتقدمة</option>
          <option value="not-related-strategic">غير مرتبط بهدف استراتيجي</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">الهدف الإستراتيجي من المستوى الثاني <span class="required">*</span></label>
        <select class="form-control" id="strategicGoalLevel2" name="strategicGoalLevel2" required data-required>
          <option value="">اختر الهدف الإستراتيجي من المستوى الثاني</option>
          <option value="enhance-reliability-availability">تعزيز موثوقية وتوفرية التشغيل</option>
          <option value="improve-data-quality">رفع مستوى جودة البيانات</option>
          <option value="enhance-integration">تعزيز التكامل</option>
          <option value="enable-digital-solutions">تمكين الحلول الرقمية والتقنيات المتقدمة</option>
          <option value="digitize-improve-experience">رقمنة وتحسين تجربة المستفيد</option>
          <option value="improve-systems-regulations">تحسين الأنظمة والتشريعات لتنويع مصادر الإيرادات</option>
          <option value="not-related-strategic-level2">غير مرتبط بهدف استراتيجي</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">تصنيف مقدم النشاط <span class="required">*</span></label>
        <select class="form-control" id="presenterCategory" name="presenterCategory" required data-required>
          <option value="">اختر تصنيف مقدم النشاط</option>
          <option value="external-expert">خبير خارجي</option>
          <option value="sfda-employee">موظف بالهيئة العامة للغذاء والدواء</option>
          <option value="trainee">متدرب</option>
          <option value="section-head">رئيس قسم</option>
          <option value="department-manager">مدير إدارة</option>
          <option value="executive-manager">مدير تنفيذي</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">تاريخ بداية النشاط <span class="required">*</span></label>
        <input type="date" class="form-control" id="activityStartDate" name="activityStartDate" required data-required>
      </div>
      <div class="form-group">
        <label class="form-label">تاريخ نهاية النشاط <span class="required">*</span></label>
        <input type="date" class="form-control" id="activityEndDate" name="activityEndDate" required data-required>
      </div>
      <div class="form-group">
        <label class="form-label">اسم مقدم النشاط <span class="required">*</span></label>
        <input type="text" class="form-control" id="presenterName" name="presenterName" required data-required placeholder="كتابة اسم الشخص أو الأشخاص الذي قاموا بتقديم النشاط المعرفي (يرجى كتابة الأسماء)">
      </div>
      <div class="form-group">
        <label class="form-label">اسم المسؤول عن قوائم الحضور من الفئة المستهدفة <span class="required">*</span></label>
        <input type="text" class="form-control" id="attendanceResponsible" name="attendanceResponsible" required data-required placeholder="اسم الشخص المسؤول عن أسماء الحضور">
      </div>
      <div class="form-group">
        <label class="form-label">إرفاق المستندات (اختياري)</label>
        <input type="file" class="form-control" id="attendanceDocuments" name="attendanceDocuments" multiple accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png" style="padding: 8px;">
        <small class="file-help">يمكن إرفاق ملفات PDF, Word, Excel، أو صور (حد أقصى 10 ملفات)</small>
      </div>
      <div class="form-group">
        <label class="form-label">الفئة المستهدفة <span class="required">*</span></label>
        <select class="form-control" id="targetAudienceType" name="targetAudienceType" required data-required>
          <option value="">اختر الفئة المستهدفة</option>
          <option value="داخل الهيئة">داخل الهيئة</option>
          <option value="خارج الهيئة">خارج الهيئة</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">يرجى كتابة الفئة المستهدفة <span class="required">*</span></label>
        <textarea class="form-control" id="targetAudienceDetails" name="targetAudienceDetails" required data-required placeholder="يرجى تحديد وكتابة تفاصيل الفئة المستهدفة بوضوح"></textarea>
      </div>
      <div class="form-group">
        <label class="form-label">عدد الحضور <span class="required">*</span></label>
        <input type="number" class="form-control" id="attendeeCount" name="attendeeCount" required data-required min="1" placeholder="يرجى كتابة عدد الحضور">
      </div>
      <div class="form-group">
        <label class="form-label">مدة النشاط (بالساعة) <span class="required">*</span></label>
        <input type="number" class="form-control" id="activityDuration" name="activityDuration" required data-required placeholder="يرجى تحديد مدة النشاط">
        <small class="file-help">مثال: 02:30 (ساعتان ونصف)</small>
      </div>
      <div class="form-group">
        <label class="form-label">مكان حفظ المحتوى المعرفي <span class="required">*</span></label>
        <input type="text" class="form-control" id="contentLocation" name="contentLocation" required data-required placeholder="يرجى كتابة أين تم حفظ المحتوى المعرفي">
      </div>
      <div class="form-group">
        <label class="form-label">إرفاق المستندات (اختياري)</label>
        <input type="file" class="form-control" id="contentDocuments" name="contentDocuments" multiple accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png,.ppt,.pptx" style="padding: 8px;">
        <small class="file-help">يمكن إرفاق المحتوى المعرفي أو المستندات ذات الصلة (حد أقصى 10 ملفات)</small>
      </div>
    `;
    formFieldsEl.style.display = 'block';
    submitBtn.style.display = 'block';
    const now = new Date();
    const firstDayOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const lastDayOfLastMonth  = new Date(now.getFullYear(), now.getMonth(), 0);
    const minDate = firstDayOfLastMonth.toISOString().split('T')[0];
    const maxDate = lastDayOfLastMonth.toISOString().split('T')[0];
    document.getElementById('activityStartDate').setAttribute('min', minDate);
    document.getElementById('activityStartDate').setAttribute('max', maxDate);
    document.getElementById('activityEndDate').setAttribute('min', minDate);
    document.getElementById('activityEndDate').setAttribute('max', maxDate);
    document.getElementById('activityStartDate').addEventListener('change', function () {
      const startDate = this.value;
      if (startDate) {
        const endDateInput = document.getElementById('activityEndDate');
        endDateInput.setAttribute('min', startDate);
        if (endDateInput.value && endDateInput.value < startDate) {
          endDateInput.value = '';
        }
      }
    });
    attachEventListeners();
    formFieldsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ===== تأثيرات تفاعلية + فحص الحقول =====
  function attachEventListeners() {
    const newFormControls = document.querySelectorAll('#formFields .form-control');
    newFormControls.forEach((control) => {
      control.addEventListener('focus', function () {
        this.parentElement.style.transform = 'scale(1.02)';
        this.parentElement.style.transition = 'transform 0.3s ease';
        const existingError = this.parentElement.querySelector('.error-message');
        if (existingError) existingError.remove();
        this.style.borderColor = '#e0e0e0';
        this.style.boxShadow = 'none';
      });
      control.addEventListener('input', function () {
        if (this.hasAttribute('required')) {
          let isValid = false;
          if (this.type === 'select-one') isValid = this.value !== '';
          else if (this.type === 'number') isValid = this.value !== '' && Number(this.value) > 0;
          else isValid = this.value.trim() !== '';
          if (isValid) {
            this.style.borderColor = '#4caf50';
            this.style.boxShadow = '0 0 0 3px rgba(76, 175, 80, 0.2)';
            const existingError = this.parentElement.querySelector('.error-message');
            if (existingError) existingError.remove();
          }
        }
      });
      control.addEventListener('blur', function () {
        this.parentElement.style.transform = 'scale(1)';
        if (this.hasAttribute('required')) {
          let isEmpty = false;
          if (this.type === 'select-one') isEmpty = this.value === '';
          else if (this.type === 'number') isEmpty = this.value === '' || Number(this.value) <= 0;
          else isEmpty = this.value.trim() === '';
          if (isEmpty) {
            this.style.borderColor = '#f44336';
            this.style.boxShadow = '0 0 0 3px rgba(244, 67, 54, 0.2)';
            const existingError = this.parentElement.querySelector('.error-message');
            if (!existingError) {
              const errorMessage = document.createElement('div');
              errorMessage.className = 'error-message';
              errorMessage.textContent = 'هذا الحقل مطلوب ولا يمكن تركه فارغاً';
              this.parentElement.appendChild(errorMessage);
            }
          }
        }
      });
    });
  }

  // ===== إرسال النموذج =====
 // ===== إرسال النموذج (النسخة النهائية والمصححة بالكامل) =====
  knowledgeForm.addEventListener('submit', async function (e) {
    e.preventDefault();
    document.getElementById('successMessage').style.display = 'none';
    document.getElementById('loading').style.display = 'block';
    
    const formData = new FormData(this);

    // --- تحويل قيم القوائم المنسدلة من ID/Value إلى نص عربي ---

    // 1. حقول الهيكل التنظيمي
    const sectorText = sectorEl.options[sectorEl.selectedIndex].text;
    formData.set('sector', sectorText);
    const departmentText = departmentEl.options[departmentEl.selectedIndex].text;
    formData.set('department', departmentText);
    const divisionText = divisionEl.options[divisionEl.selectedIndex].text;
    formData.set('division', divisionText);
    const sectionText = sectionEl.options[sectionEl.selectedIndex].text;
    formData.set('section', sectionText);
    
    if (document.getElementById('activityType')) {
        const activityTypeText = document.getElementById('activityType').options[document.getElementById('activityType').selectedIndex].text;
        formData.set('activityType', activityTypeText);
    }
    if (document.getElementById('strategicGoalLevel1')) {
        const goal1Text = document.getElementById('strategicGoalLevel1').options[document.getElementById('strategicGoalLevel1').selectedIndex].text;
        formData.set('strategicGoalLevel1', goal1Text);
    }
    if (document.getElementById('strategicGoalLevel2')) {
        const goal2Text = document.getElementById('strategicGoalLevel2').options[document.getElementById('strategicGoalLevel2').selectedIndex].text;
        formData.set('strategicGoalLevel2', goal2Text);
    }
    if (document.getElementById('presenterCategory')) {
        const presenterCategoryText = document.getElementById('presenterCategory').options[document.getElementById('presenterCategory').selectedIndex].text;
        formData.set('presenterCategory', presenterCategoryText);
    }

    try {
        const response = await fetch('/form', {
            method: 'POST',
            body: formData,
        });
        document.getElementById('loading').style.display = 'none';
        if (response.ok) {
            const resultText = await response.text();
            const successDiv = document.getElementById('successMessage');
            successDiv.innerHTML = `<h3>${resultText}</h3>`;
            successDiv.style.display = 'block';
            successDiv.scrollIntoView({ behavior: 'smooth' });
            setTimeout(async function () {
                knowledgeForm.reset();
                successDiv.style.display = 'none';
                await initOrgLists();
                activityQuestionEl.style.display = 'none';
                formFieldsEl.style.display = 'none';
                submitBtn.style.display = 'none';
                noActivitiesMessage.style.display = 'none';
                window.scrollTo({ top: 0, behavior: 'smooth' });
            }, 3000);
        } else {
            console.error('❌ فشل في حفظ الرد');
            alert('حدث خطأ أثناء إرسال النموذج. الرجاء المحاولة مجدداً.');
        }
    } catch (error) {
        document.getElementById('loading').style.display = 'none';
        console.error('⚠️ خطأ في الاتصال بالسيرفر:', error);
        alert('فشل في الاتصال بخادم حفظ البيانات.');
    }
  });

  // ===== تشغيل أولي =====
  initOrgLists();
});