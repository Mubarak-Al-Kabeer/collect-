/**
 * دفتر القطية — Google Sheets + Apps Script
 * الشيت: إنت بس اللي تعدّل فيه (الأسماء، الترتيب، الدفعات، التسليم).
 * صفحة الويب: للكل، عرض فقط، بدون تسجيل دخول.
 *
 * هذا الملف نسخة من الكود اللي في Apps Script. أي تعديل هنا لازم تنسخه هناك
 * وبعدين: Deploy ← Manage deployments ← Edit ← Version: New version.
 */

const MARKER = 'قطية';          // A1 في كل تبويب قطية
const FIRST_ROW = 7;            // أول صف للأسماء
const MAX = 10;                 // أقصى عدد أشخاص
const MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];

// نفس المنطقة الزمنية حقت الشيت، عشان التواريخ ما تنزاح يوم
function tz() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }

/* ---------- القائمة داخل الشيت ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🪙 القطية')
    .addItem('إنشاء قطية جديدة', 'createQatiya')
    .addItem('رابط صفحة العرض', 'showLink')
    .addToUi();
}

function createQatiya() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('قطية جديدة', 'اكتب اسم القطية:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const name = res.getResponseText().trim() || 'قطية';
  const ss = SpreadsheetApp.getActive();
  let tabName = name, i = 2;
  while (ss.getSheetByName(tabName)) tabName = name + ' ' + (i++);
  const sh = ss.insertSheet(tabName);
  sh.setRightToLeft(true);

  // الإعدادات
  sh.getRange('A1:B4').setValues([
    [MARKER, name],
    ['المبلغ الشهري (د.ك)', 50],
    ['تاريخ أول استحقاق', new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1)],
    ['عدد الأشخاص', '=COUNTA(B7:B16)']
  ]);
  sh.getRange('A1:A4').setFontWeight('bold').setBackground('#e2efe9');
  sh.getRange('B1').setFontWeight('bold').setFontSize(13);
  sh.getRange('B3').setNumberFormat('d mmmm yyyy');
  sh.getRange('A5').setValue('رتّب الأسماء حسب دور الاستلام. علّم ✓ تحت "دفع دور X" لكل شخص دفع، و ✓ في "استلم" لما تسلّم المبلغ.')
    .setFontColor('#5d6b64').setFontSize(9);

  // الجدول
  const headers = ['الدور', 'الاسم', 'تاريخ الاستحقاق', 'استلم'];
  for (let k = 1; k <= MAX; k++) headers.push('دفع دور ' + k);
  sh.getRange(6, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0f6b55').setFontColor('#ffffff').setHorizontalAlignment('center');

  const rows = [];
  for (let r = 0; r < MAX; r++) {
    const row = FIRST_ROW + r;
    rows.push([r + 1, '', `=IF(B${row}="","",EDATE($B$3,A${row}-1))`]);
  }
  sh.getRange(FIRST_ROW, 1, MAX, 3).setValues(rows);
  sh.getRange(FIRST_ROW, 3, MAX, 1).setNumberFormat('d mmmm yyyy');
  sh.getRange(FIRST_ROW, 4, MAX, 1 + MAX).insertCheckboxes();
  sh.getRange(FIRST_ROW, 1, MAX, 1).setHorizontalAlignment('center').setFontColor('#a5761f').setFontWeight('bold');

  // تلوين: اللي استلم أخضر
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$D${FIRST_ROW}=TRUE`).setBackground('#e1f2e8')
    .setRanges([sh.getRange(FIRST_ROW, 1, MAX, 4)]).build();
  sh.setConditionalFormatRules([rule]);

  sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 140);
  sh.setFrozenRows(6);
  sh.getRange('B7').activate();
  ui.alert('تم إنشاء "' + tabName + '". اكتب الأسماء من B7 حسب ترتيب الاستلام، وعدّل المبلغ والتاريخ فوق.');
}

function showLink() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  ui.alert(url ? 'رابط صفحة العرض:\n' + url
               : 'ما في رابط للحين. سوّ: Deploy ← New deployment ← Web app، وخل Who has access = Anyone.');
}

/* ---------- قراءة البيانات ---------- */
function fmtDate(d, zone) {
  if (!(d instanceof Date)) return '';
  const day = Number(Utilities.formatDate(d, zone, 'd'));
  const m = Number(Utilities.formatDate(d, zone, 'M'));
  const y = Utilities.formatDate(d, zone, 'yyyy');
  return day + ' ' + MONTHS[m - 1] + ' ' + y;
}

function getAll() {
  const zone = tz();
  const out = [];
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    if (String(sh.getRange('A1').getValue()).trim() !== MARKER) return;
    const top = sh.getRange('B1:B3').getValues();
    const grid = sh.getRange(FIRST_ROW, 1, MAX, 4 + MAX).getValues();
    // رقم الصف الأصلي لكل اسم (عشان لو في صف فاضي بالنص ما تنزاح أعمدة الدفع)
    const rows = grid.map((r, i) => i).filter(i => String(grid[i][1]).trim() !== '');
    const name = String(top[0][0] || sh.getName());
    const amount = Number(top[1][0]) || 0;
    out.push({
      name, amount,
      members: rows.map(i => {
        const r = grid[i];
        return {
          name: String(r[1]).trim(),
          due: fmtDate(r[2], zone),
          received: r[3] === true,
          paid: rows.map(j => r[4 + j] === true)   // paid[k] = دفع في دور k (عمود "دفع دور" حق صاحب الدور k)
        };
      })
    });
  });
  return { groups: out, updated: Utilities.formatDate(new Date(), zone, 'd/M/yyyy h:mm a') };
}

/* ---------- السيرفر (API) — الواجهة على GitHub Pages تقرأ منه ---------- */
function doGet() {
  let out;
  try { out = getAll(); } catch (err) { out = { error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
