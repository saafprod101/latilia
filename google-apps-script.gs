/**
 * لتيليا | Latilia — استقبال طلبات صفحة الهبوط في Google Sheets
 * الشيت: طلبات لتيليا - Latilia Orders
 * مشروع Apps Script: Latilia Orders Webhook (منشور كـ Web App: Execute as Me، Access: Anyone)
 *
 * عند تعديل الكود: Deploy ← Manage deployments ← تعديل (القلم) ← Version: New version ← Deploy
 * حتى يبقى نفس الرابط الموجود في index.html.
 */
const SHEET_ID = '1xUu3Bi93emP5VmXz1dWllvPvhnmZzBxwHOig81ldozs';
const HEADERS = ['رقم الطلب', 'التاريخ والوقت', 'الاسم', 'الجوال', 'الحي', 'طريقة الدفع', 'الأطباق', 'عدد الأطباق', 'الإجمالي', 'ملاحظات', 'المصدر', 'الحالة'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return json_({ result: 'busy' });
  }
  try {
    var sheet = getSheet_();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      safe_(data.orderId),
      Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'),
      safe_(data.name),
      "'" + String(data.phone || '').replace(/\D/g, ''),   // لحفظ الصفر الأول في رقم الجوال
      safe_(data.district),
      safe_(data.payment || data.address),   // طريقة الدفع: كاش / شبكة
      safe_(data.itemsText),
      Number(data.itemsCount || 0),
      Number(data.total || 0) + ' ر.س',
      safe_(data.notes || 'لا يوجد'),
      safe_(data.source),
      'جديد'
    ]);
    return json_({ result: 'success', orderId: data.orderId || null });
  } catch (error) {
    return json_({ result: 'error', error: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// افتح رابط /exec في المتصفح للتأكد أن الربط يعمل
function doGet() {
  return ContentService.createTextOutput('Latilia orders webhook is running');
}

function getSheet_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  // تحديث عنوان العمود القديم «العنوان» إلى «طريقة الدفع»
  if (sheet.getLastRow() > 0 && sheet.getRange(1, 6).getValue() === 'العنوان') {
    sheet.getRange(1, 6).setValue('طريقة الدفع');
  }
  if (sheet.getLastRow() === 0) {
    sheet.setName('الطلبات');
    sheet.appendRow(HEADERS);
    sheet.setRightToLeft(true);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#606D39').setFontColor('#FFF0D8').setHorizontalAlignment('center');
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 280);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['جديد', 'قيد التحضير', 'خرج للتوصيل', 'تم التسليم', 'ملغي'], true).build();
    sheet.getRange(2, HEADERS.length, 2000, 1).setDataValidation(rule);
  }
  return sheet;
}

// يمنع تحويل النص إلى معادلة داخل الشيت
function safe_(v) {
  var s = String(v == null ? '' : v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// تشغيل يدوي لإنشاء العناوين: اختر setup ثم Run
function setup() {
  getSheet_();
}
