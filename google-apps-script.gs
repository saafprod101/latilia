/**
 * لتيليا | Latilia — استقبال طلبات صفحة الهبوط في Google Sheets
 * ---------------------------------------------------------------
 * طريقة التركيب:
 * 1) افتح Google Sheet جديد ← الإضافات (Extensions) ← Apps Script
 * 2) احذف الكود الموجود والصق هذا الملف كاملاً ← حفظ
 * 3) نشر (Deploy) ← نشر جديد (New deployment) ← النوع: تطبيق ويب (Web app)
 *      - التنفيذ باسم (Execute as): أنا (Me)
 *      - من لديه صلاحية الوصول (Who has access): أي شخص (Anyone)
 * 4) وافق على الصلاحيات، ثم انسخ رابط Web app (ينتهي بـ /exec)
 * 5) ضعه في index.html داخل: const GOOGLE_SHEET_URL = "الرابط";
 *
 * ملاحظة: عند تعديل هذا الكود لاحقاً ← Deploy ← Manage deployments ← تعديل ← Version: New version
 * حتى يبقى نفس الرابط ويعمل التحديث.
 */

const SHEET_NAME = 'الطلبات';
const TIMEZONE = 'Asia/Riyadh';
const HEADERS = [
  'رقم الطلب', 'التاريخ والوقت', 'الاسم', 'الجوال', 'الحي', 'العنوان التفصيلي',
  'الأطباق', 'عدد الأطباق', 'الإجمالي (ر.س)', 'ملاحظات', 'المصدر', 'الحملة', 'الحالة'
];
const STATUSES = ['جديد', 'قيد التحضير', 'خرج للتوصيل', 'تم التسليم', 'ملغي'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // يمنع تداخل الطلبات المتزامنة
  try {
    const raw = (e && e.postData && e.postData.contents) || '{}';
    const data = JSON.parse(raw);
    const sheet = getSheet_();

    const items = Array.isArray(data.items) ? data.items : [];
    const itemsText = items.length
      ? items.map(function (i) { return i.qty + 'x ' + i.name + ' (' + (i.qty * i.price) + ' ر.س)'; }).join('\n')
      : (data.itemsText || '');
    const itemsCount = items.reduce(function (s, i) { return s + Number(i.qty || 0); }, 0) || Number(data.itemsCount || 0);
    const total = items.reduce(function (s, i) { return s + Number(i.qty || 0) * Number(i.price || 0); }, 0) || Number(data.total || 0);

    sheet.appendRow([
      safe_(data.orderId),
      Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'),
      safe_(data.name),
      "'" + String(data.phone || '').replace(/\D/g, ''), // للحفاظ على الصفر في بداية الرقم
      safe_(data.district),
      safe_(data.address),
      safe_(itemsText),
      itemsCount,
      total,
      safe_(data.notes),
      safe_(data.source),
      safe_(data.campaign),
      'جديد'
    ]);

    return json_({ ok: true, orderId: data.orderId || null });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// للتأكد أن الرابط يعمل: افتح رابط /exec في المتصفح
function doGet() {
  return ContentService.createTextOutput('✅ Latilia orders webhook is running');
}

/* ---------------- Helpers ---------------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setRightToLeft(true);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#5D6D4E')
      .setFontColor('#FBF9F5')
      .setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(6, 240);
    sheet.setColumnWidth(7, 260);
    sheet.setColumnWidth(10, 200);
    sheet.getRange('G:G').setWrap(true);

    // قائمة منسدلة لحالة الطلب
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
    sheet.getRange(2, HEADERS.length, 1000, 1).setDataValidation(rule);
  }
  return sheet;
}

// يمنع حقن المعادلات (Formula Injection) في الشيت
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// اختبار يدوي من محرر Apps Script: اختر testDoPost ثم Run
function testDoPost() {
  const res = doPost({
    postData: {
      contents: JSON.stringify({
        orderId: 'LT-TEST-0001',
        name: 'تجربة',
        phone: '0500000000',
        district: 'المزروع',
        address: 'شارع تجريبي',
        notes: 'طلب اختبار',
        items: [{ name: 'باستا اللومي الحساوي', qty: 2, price: 25 }],
        source: 'test'
      })
    }
  });
  Logger.log(res.getContent());
}
