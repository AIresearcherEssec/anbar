/**
 * انبار مغازه — بک‌اند گوگل‌شیت (نسخهٔ ۵)
 * ------------------------------------------------------------
 * دو شیت دارد:
 *   «کالاها»      → موجودی زندهٔ همین لحظه؛ هر کالا فقط یک سطر.
 *   «تراکنش‌ها»   → دفتر رویدادها؛ هر کم و زیاد شدن یک سطر.
 *
 * نصب: Extensions > Apps Script > کل کد قبلی را پاک کن و این را پیست کن >
 *      ذخیره > Deploy > Manage deployments > مداد > New version > Deploy
 */

const TOKEN = 'Hadis1376';

const SH_ITEMS  = 'کالاها';
const SH_LOG    = 'تراکنش‌ها';
const HDR_ITEMS = ['id', 'بارکد', 'نام کالا', 'قیمت', 'موجودی', 'واحد', 'آخرین تغییر',
                   'دسته‌بندی', 'زیردسته', 'زود انقضا', 'تاریخ انقضا'];
const HDR_LOG   = ['زمان', 'نام کالا', 'تغییر', 'موجودی بعد', 'دستگاه'];
const PROP_OPS  = 'recentOpIds';

/* ---------------- ورودی‌ها ---------------- */

const VERSION = 5;   // با باز کردن آدرس /exec در مرورگر دیده می‌شود

function doGet() {
  return json({ ok: true, v: VERSION, cols: HDR_ITEMS.length,
                msg: 'API انبار فعال است. نسخهٔ ' + VERSION + ' با ' + HDR_ITEMS.length + ' ستون.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) return json({ ok: false, error: 'توکن نامعتبر است' });

    lock.waitLock(30000);          // اگر گوشی دیگری در حال نوشتن است، نوبت می‌گیرد

    setup_();
    ensureIds_();                  // سطرهایی که دستی در شیت اضافه شده‌اند id می‌گیرند
    var applied = applyOps_(req.ops || []);
    return json({ ok: true, applied: applied, items: readItems_(), t: Date.now() });

  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- ساخت شیت‌ها ---------------- */

function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var it = ss.getSheetByName(SH_ITEMS);
  if (!it) {
    it = ss.insertSheet(SH_ITEMS);
    it.getRange(1, 1, 1, HDR_ITEMS.length).setValues([HDR_ITEMS])
      .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    it.setFrozenRows(1);
    it.setColumnWidth(2, 150);
    it.setColumnWidth(3, 300);
    it.setColumnWidth(8, 150);
    it.hideColumns(1);            // ستون id فنی است
  }

  // اگر شیت از قبل بوده و ستون‌های تازه سرستون ندارند، همین‌جا نوشته می‌شوند
  ensureHeader_(it);

  var lg = ss.getSheetByName(SH_LOG);
  if (!lg) {
    lg = ss.insertSheet(SH_LOG);
    lg.getRange(1, 1, 1, HDR_LOG.length).setValues([HDR_LOG])
      .setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
    lg.setFrozenRows(1);
    lg.setColumnWidth(1, 170);
    lg.setColumnWidth(2, 300);
    lg.getRange('A:A').setNumberFormat('yyyy/mm/dd  HH:mm');   // ساعت هم دیده شود
  }

  var s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);

  return ss;
}

/* سرستون‌های شیت را با نسخهٔ فعلی هماهنگ می‌کند */
function ensureHeader_(sh) {
  var cur = sh.getRange(1, 1, 1, HDR_ITEMS.length).getValues()[0];
  var need = false;
  for (var i = 0; i < HDR_ITEMS.length; i++) {
    if (String(cur[i] || '').trim() !== HDR_ITEMS[i]) { need = true; break; }
  }
  if (!need) return;
  sh.getRange(1, 1, 1, HDR_ITEMS.length).setValues([HDR_ITEMS])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  sh.getRange(1, HDR_ITEMS.length, sh.getMaxRows(), 1).setNumberFormat('@'); // انقضا متن بماند
}

/* ---------------- شناسهٔ خودکار برای سطرهای دستی ---------------- */
/* اگر کسی مستقیم در شیت سطر اضافه کند (مثلاً با پیست کردن از فاکتور)،
   ستون id خالی می‌ماند. اینجا برایشان شناسهٔ یکتا ساخته و نوشته می‌شود. */

function ensureIds_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ITEMS);
  var last = sh.getLastRow();
  if (last < 2) return;

  var n    = last - 1;
  var ids  = sh.getRange(2, 1, n, 1).getValues();
  var name = sh.getRange(2, 3, n, 1).getValues();
  var changed = false;

  for (var i = 0; i < n; i++) {
    if (!ids[i][0] && String(name[i][0] || '').trim()) {
      ids[i][0] = 'g' + Date.now().toString(36) + i.toString(36) +
                  Math.floor(Math.random() * 1296).toString(36);
      changed = true;
    }
  }
  if (changed) sh.getRange(2, 1, n, 1).setValues(ids);
}

/* تاریخ انقضا به‌صورت متن شمسی «۱۴۰۵/۰۹/۰۳» نگه داشته می‌شود.
   اگر گوگل‌شیت آن را به تاریخ تبدیل کرد، دوباره به متن برمی‌گردانیم. */
function fmtExp_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  return String(v).trim();
}

/* ---------------- خواندن ---------------- */

function readItems_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ITEMS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HDR_ITEMS.length).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && !r[2]) continue;
    out.push({
      id:    String(r[0] || ('row' + (i + 2))),
      code:  r[1] === '' || r[1] == null ? '' : String(r[1]),
      name:  String(r[2] || ''),
      price: Number(r[3]) || 0,
      stock: Number(r[4]) || 0,
      unit:  String(r[5] || 'عدد'),
      at:    r[6] ? new Date(r[6]).getTime() : 0,
      grp:   String(r[7]  == null ? '' : r[7]).trim(),
      cat:   String(r[8]  == null ? '' : r[8]).trim(),
      per:   String(r[9]  == null ? '' : r[9]).trim() === 'بله',
      exp:   fmtExp_(r[10])
    });
  }
  return out;
}

function indexRows_(sh) {
  var last = sh.getLastRow();
  var map = {};
  if (last < 2) return map;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0]) map[String(ids[i][0])] = i + 2;
  }
  return map;
}

/* ---------------- جلوگیری از ثبت تکراری ---------------- */
/* شناسهٔ عملیات‌ها در حافظهٔ خود اسکریپت نگه داشته می‌شود، نه در شیت،
   تا دفتر تراکنش‌ها فقط ستون‌های قابل‌فهم داشته باشد. */

function loadOpIds_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PROP_OPS);
  var seen = {};
  if (raw) {
    try {
      var arr = JSON.parse(raw);
      for (var i = 0; i < arr.length; i++) seen[arr[i]] = true;
    } catch (e) {}
  }
  return seen;
}

function saveOpIds_(seen) {
  var arr = Object.keys(seen);
  if (arr.length > 400) arr = arr.slice(arr.length - 400);
  PropertiesService.getScriptProperties().setProperty(PROP_OPS, JSON.stringify(arr));
}

/* ---------------- اعمال تغییرات ---------------- */

function applyOps_(ops) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var it  = ss.getSheetByName(SH_ITEMS);
  var lg  = ss.getSheetByName(SH_LOG);
  var map = indexRows_(it);
  var seen = loadOpIds_();
  var applied = [];
  var logRows = [];
  var touched = false;

  for (var i = 0; i < ops.length; i++) {
    var op = ops[i];
    if (!op || !op.opId) continue;
    if (seen[op.opId]) { applied.push(op.opId); continue; }

    try {
      if (op.op === 'delta') {
        var row = map[op.id];
        if (!row) { applied.push(op.opId); seen[op.opId] = true; continue; }
        var cur   = Number(it.getRange(row, 5).getValue()) || 0;
        var after = cur + Number(op.qty || 0);
        it.getRange(row, 5).setValue(after);
        it.getRange(row, 7).setValue(new Date());
        logRows.push([new Date(), it.getRange(row, 3).getValue(),
                      Number(op.qty || 0), after, op.device || '']);

      } else if (op.op === 'upsert') {
        var r2 = map[op.id];
        var vals = [op.id, op.code || '', op.name || '', Number(op.price) || 0,
                    Number(op.stock) || 0, op.unit || 'عدد', new Date(),
                    op.grp || '', op.cat || '', op.per ? 'بله' : 'خیر', op.exp || ''];
        if (r2) {
          it.getRange(r2, 1, 1, HDR_ITEMS.length).setValues([vals]);
        } else {
          it.appendRow(vals);
          map[op.id] = it.getLastRow();
          logRows.push([new Date(), op.name || '', Number(op.stock) || 0,
                        Number(op.stock) || 0, op.device || '']);
        }

      } else if (op.op === 'meta') {
        // نام و قیمت را عوض می‌کند ولی به موجودی دست نمی‌زند
        var r4 = map[op.id];
        if (r4) {
          it.getRange(r4, 2).setValue(op.code || '');
          it.getRange(r4, 3).setValue(op.name || '');
          it.getRange(r4, 4).setValue(Number(op.price) || 0);
          it.getRange(r4, 6).setValue(op.unit || 'عدد');
          it.getRange(r4, 7).setValue(new Date());
          it.getRange(r4, 8).setValue(op.grp || '');
          it.getRange(r4, 9).setValue(op.cat || '');
          it.getRange(r4, 10).setValue(op.per ? 'بله' : 'خیر');
          it.getRange(r4, 11).setValue(op.exp || '');
        }

      } else if (op.op === 'delete') {
        var r3 = map[op.id];
        if (r3) { it.deleteRow(r3); map = indexRows_(it); }
      }

      applied.push(op.opId);
      seen[op.opId] = true;
      touched = true;

    } catch (err) {
      // ناموفق: در صف گوشی می‌ماند و دفعهٔ بعد دوباره تلاش می‌شود
    }
  }

  if (logRows.length) {
    lg.getRange(lg.getLastRow() + 1, 1, logRows.length, HDR_LOG.length).setValues(logRows);
  }
  if (touched) saveOpIds_(seen);
  return applied;
}

/* ---------------- منوی کمکی داخل شیت ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('انبار')
    .addItem('ساخت شیت‌های لازم', 'setup_')
    .addToUi();
}
