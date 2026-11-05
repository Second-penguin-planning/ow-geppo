/**
 * ワンワールド月報告アプリ（出勤・経費をスマホで記入）の記録を、このスプレッドシートに貯める。
 * 「月報告ひな形.xlsx」を Google スプレッドシートに変えたファイルの「拡張機能 → Apps Script」に貼り付け、
 * ウェブアプリとして公開して使う。
 *
 * 記録のシート（なければ自動で作る）
 *   出勤データ：1日1行　　交通費データ：1区間1行（同じ旅IDで1回の移動）　　立替データ：1件1行
 *   報告履歴：PDFを作るたびに1行
 * 領収書の写真は、このスプレッドシートと同じフォルダの「領収書」フォルダに入る
 * 報告のシート（ひな形・PDFに出す）
 *   出勤簿／立替経費精算書／交通費精算書
 *
 * 月度は「前の月の21日〜その月の20日」。11月度 = 10/21〜11/20。
 */
const TZ = 'Asia/Tokyo';
const ATT = { name: '出勤データ', headers: ['日付', '区分', '出勤', '退勤', '休憩', '実働', '備考', '更新日時'] };
const TRIP = { name: '交通費データ', headers: ['ID', '旅ID', '日付', '目的', '出発', '到着', '交通機関', '金額', '順番'] };
const EXP = { name: '立替データ', headers: ['ID', '日付', '内容', '支払先', '領収書', '金額', '写真'] };
const REP = { name: '報告履歴', headers: ['月度', '作成日時', '申請日', '出勤日数', '実働', '立替合計', 'PDF'] };
const DATA_SHEETS = [ATT, TRIP, EXP, REP];
const SHEET_ATT = '出勤簿';
const SHEET_EXP = '立替経費精算書';
const SHEET_TRIP = '交通費精算書';
const TEXT_COLS = /日付|出勤|退勤|休憩|実働|月度/;

function sheetOf_(def) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(def.name);
  if (!sheet) {
    sheet = ss.insertSheet(def.name);
    sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < def.headers.length) {
    // あとから足した列の見出しを入れる
    sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 日付・時刻は先頭に ' を付けて文字のまま書く
function cell_(v, h) {
  return TEXT_COLS.test(h) && v !== '' && v != null ? "'" + v : (v == null ? '' : v);
}

function readAll_(def) {
  const sheet = sheetOf_(def);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const shown = sheet.getRange(2, 1, last - 1, def.headers.length).getDisplayValues();
  const values = sheet.getRange(2, 1, last - 1, def.headers.length).getValues();
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    def.headers.forEach((h, j) => { o[h] = TEXT_COLS.test(h) ? shown[i][j] : r[j]; });
    return o;
  });
}

// key列の値が一致する行の番号（下から探す）。全部ほしいときは all=true
function rowsOf_(def, key, val, all) {
  const sheet = sheetOf_(def);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const col = def.headers.indexOf(key) + 1;
  const vals = sheet.getRange(2, col, last - 1, 1).getDisplayValues();
  const out = [];
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]) === String(val)) { out.push(i + 2); if (!all) break; }
  }
  return out;  // 下から順（そのまま deleteRow してよい）
}

function upsert_(def, key, row) {
  const sheet = sheetOf_(def);
  const values = [def.headers.map(h => cell_(row[h], h))];
  const found = rowsOf_(def, key, row[key]);
  if (found.length) sheet.getRange(found[0], 1, 1, def.headers.length).setValues(values);
  else sheet.appendRow(values[0]);
}

function remove_(def, key, val) {
  const sheet = sheetOf_(def);
  rowsOf_(def, key, val, true).forEach(r => sheet.deleteRow(r));
}

function doPost(e) {
  // 領収書の読み取りは時間がかかるので、記録の書き込みとは別に（ロックの外で）行う
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'receipt') return json_(Object.assign({ ok: true }, readReceipt_(req)));
    if (req.action === 'receiptDiscard') { discardReceipt_(req.photoId); return json_({ ok: true }); }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);

    if (req.action === 'attend') {
      // 区分も時刻もない日は消す
      const r = req.row;
      if (!r['区分'] && !r['出勤'] && !r['退勤'] && !r['備考']) remove_(ATT, '日付', r['日付']);
      else upsert_(ATT, '日付', Object.assign({}, r, { '更新日時': Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') }));
      return json_({ ok: true });
    }

    if (req.action === 'trip') {
      // 1回の移動（旅ID）の区間を全部入れ替える。legs が空なら消す
      remove_(TRIP, '旅ID', req.tripId);
      const sheet = sheetOf_(TRIP);
      (req.legs || []).forEach((leg, i) => {
        sheet.appendRow(TRIP.headers.map(h => cell_(h === '順番' ? i + 1 : h === '旅ID' ? req.tripId : leg[h], h)));
      });
      return json_({ ok: true });
    }

    if (req.action === 'expense') {
      if (req.remove) remove_(EXP, 'ID', req.row['ID']);
      else { upsert_(EXP, 'ID', req.row); nameReceipt_(req.row); }
      return json_({ ok: true });
    }

    if (req.action === 'report') {
      return json_(Object.assign({ ok: true }, makeReport_(req)));
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// アプリを開いたときに全部返す（1年でも数百行なので全部でよい）
function doGet() {
  const strip = rows => rows.map(r => { delete r._row; return r; });
  return json_({
    ok: true,
    attend: strip(readAll_(ATT)),
    trips: strip(readAll_(TRIP)),
    expenses: strip(readAll_(EXP)),
    reports: strip(readAll_(REP)),
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  });
}

// ---- 報告（PDF） ----

function ymd_(s) {  // '2026/11/03' → [2026, 11, 3]
  const m = String(s).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  return m ? [Number(m[1]), Number(m[2]), Number(m[3])] : null;
}

// 月度（'2026-11'）の期間に入るか
function inPeriod_(dateStr, y, m) {
  const d = ymd_(dateStr);
  if (!d) return false;
  const key = d[0] * 10000 + d[1] * 100 + d[2];
  const py = m === 1 ? y - 1 : y, pm = m === 1 ? 12 : m - 1;
  return key >= py * 10000 + pm * 100 + 21 && key <= y * 10000 + m * 100 + 20;
}

function sortKey_(s) { const d = ymd_(s); return d ? d[0] * 10000 + d[1] * 100 + d[2] : 0; }

function makeReport_(req) {
  const [y, m] = req.month.split('-').map(Number);
  const applied = req.applyDate || `${y}/${String(m).padStart(2, '0')}/20`;
  const ad = ymd_(applied);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAtt = ss.getSheetByName(SHEET_ATT), shExp = ss.getSheetByName(SHEET_EXP), shTrip = ss.getSheetByName(SHEET_TRIP);
  if (!shAtt || !shExp || !shTrip) throw new Error('ひな形のシート（出勤簿・立替経費精算書・交通費精算書）が見つかりません');

  // 出勤簿：8〜18行目＝前月21〜31日、19〜38行目＝当月1〜20日
  shAtt.getRange('B2').setValue(y);
  shAtt.getRange('G2').setValue(m);
  shAtt.getRange('F8:I38').clearContent();
  shAtt.getRange('K8:N38').clearContent();
  const attend = readAll_(ATT).filter(r => inPeriod_(r['日付'], y, m));
  let days = 0, minutes = 0;
  attend.forEach(r => {
    const d = ymd_(r['日付']);
    const row = d[2] >= 21 ? 8 + d[2] - 21 : 18 + d[2];
    shAtt.getRange(row, 6, 1, 4).setValues([[r['区分'] || '', r['出勤'] || '', r['退勤'] || '', r['休憩'] || '']]);
    if (r['備考']) shAtt.getRange(row, 14).setValue(r['備考']);
    if (r['区分'] === '出勤') days++;
    const w = String(r['実働']).match(/(\d+):(\d+)/);
    if (w) minutes += Number(w[1]) * 60 + Number(w[2]);
  });

  // 交通費精算書：7〜21行目（15行）。日付・目的は旅の最初の区間だけ
  shTrip.getRange('L1').setValue(ad[0]); shTrip.getRange('N1').setValue(ad[1]); shTrip.getRange('P1').setValue(ad[2]);
  shTrip.getRange('A4').setValue(y); shTrip.getRange('C4').setValue(m);
  ['A', 'C', 'F', 'I', 'L', 'N'].forEach(c => shTrip.getRange(`${c}7:${c}21`).clearContent());
  const legs = readAll_(TRIP).filter(r => inPeriod_(r['日付'], y, m))
    .sort((a, b) => sortKey_(a['日付']) - sortKey_(b['日付']) || String(a['旅ID']).localeCompare(String(b['旅ID'])) || a['順番'] - b['順番']);
  if (legs.length > 15) throw new Error(`交通費の区間が ${legs.length} 行あり、精算書の15行に入りません`);
  let tripTotal = 0, lastTripDate = '', prevTrip = null;
  legs.forEach((leg, i) => {
    const row = 7 + i;
    const first = leg['旅ID'] !== prevTrip;
    prevTrip = leg['旅ID'];
    if (first) { shTrip.getRange(row, 1).setValue(leg['日付']); shTrip.getRange(row, 3).setValue(leg['目的']); }
    shTrip.getRange(row, 6).setValue(leg['出発']);
    shTrip.getRange(row, 9).setValue(leg['到着']);
    shTrip.getRange(row, 12).setValue(leg['交通機関']);
    shTrip.getRange(row, 14).setValue(Number(leg['金額']) || 0);
    tripTotal += Number(leg['金額']) || 0;
    lastTripDate = leg['日付'];
  });

  // 立替経費精算書：8〜29行目（22行）。1行目に交通費の合計
  shExp.getRange('B5').setValue(applied);
  shExp.getRange('A8:E29').clearContent();
  const rows = [];
  if (tripTotal > 0) rows.push([lastTripDate, req.tripLabel || '出張 交通費（電車）清算', req.tripPayee || '', '無', tripTotal]);
  readAll_(EXP).filter(r => inPeriod_(r['日付'], y, m))
    .sort((a, b) => sortKey_(a['日付']) - sortKey_(b['日付']))
    .forEach(r => rows.push([r['日付'], r['内容'], r['支払先'], r['領収書'], Number(r['金額']) || 0]));
  if (rows.length > 22) throw new Error(`立替が ${rows.length} 件あり、精算書の22行に入りません`);
  if (rows.length) shExp.getRange(8, 1, rows.length, 5).setValues(rows);
  const expTotal = rows.reduce((s, r) => s + r[4], 0);

  // PDF：記録のシート（と、交通費がない月は交通費精算書）を隠して、ブック全体を書き出す
  const hide = DATA_SHEETS.map(d => ss.getSheetByName(d.name)).filter(Boolean);
  if (!legs.length) hide.push(shTrip);
  ss.getSheets().forEach(s => { if ([SHEET_ATT, SHEET_EXP, SHEET_TRIP].indexOf(s.getName()) === -1 && hide.indexOf(s) === -1) hide.push(s); });
  hide.forEach(s => s.hideSheet());
  SpreadsheetApp.flush();
  let blob;
  try {
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&size=A4&portrait=true` +
      '&scale=4&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&fzr=false' +
      '&top_margin=0.6&bottom_margin=0.6&left_margin=0.6&right_margin=0.6&horizontal_alignment=CENTER';
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    blob = res.getBlob();
  } finally {
    hide.forEach(s => s.showSheet());
  }

  // このスプレッドシートと同じフォルダに保存。同じ名前の古いPDFはゴミ箱へ
  const yymmdd = String(ad[0]).slice(2) + String(ad[1]).padStart(2, '0') + String(ad[2]).padStart(2, '0');
  const name = `${yymmdd}_ワンワールド月報告_徳永_${y}年${m}月度.pdf`;
  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  const olds = folder.getFilesByName(name);
  while (olds.hasNext()) olds.next().setTrashed(true);
  const file = folder.createFile(blob.setName(name));

  const work = `${Math.floor(minutes / 60)}:${String(minutes % 60).padStart(2, '0')}`;
  sheetOf_(REP).appendRow([cell_(req.month, '月度'), cell_(Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'), '月度'),
    cell_(applied, '日付'), days, cell_(work, '実働'), expTotal, file.getUrl()]);
  return { pdfUrl: file.getUrl(), name, days, work, expTotal, tripTotal };
}

// ---- 領収書の写真 ----

function receiptFolder_() {
  const parent = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  const it = parent.getFoldersByName('領収書');
  return it.hasNext() ? it.next() : parent.createFolder('領収書');
}

function inReceiptFolder_(file) {
  const id = receiptFolder_().getId();
  const ps = file.getParents();
  while (ps.hasNext()) if (ps.next().getId() === id) return true;
  return false;
}

// 写真を「領収書」フォルダに保存し、Google ドキュメントに変える（Drive の文字読み取り）→ 文字を返す。
// 日付・金額・支払先の取り出しはページ側で行う（ここを直さずに直せるように）
function readReceipt_(req) {
  const mime = req.mime || 'image/jpeg';
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyMMdd_HHmmss');
  const blob = Utilities.newBlob(Utilities.base64Decode(req.image), mime, `${stamp}_領収書.jpg`);
  const photo = receiptFolder_().createFile(blob);
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const out = { text: '', photoUrl: photo.getUrl(), photoId: photo.getId() };
  // 文字が読めなくても写真は返す（手で入れて、写真付きで保存できるように）
  try {
    const res = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${photo.getId()}/copy?ocrLanguage=ja`, {
      method: 'post', contentType: 'application/json', headers: auth, muteHttpExceptions: true,
      payload: JSON.stringify({ name: stamp + '_読み取り', mimeType: 'application/vnd.google-apps.document' }),
    });
    if (res.getResponseCode() !== 200) {
      let why = '';
      try { why = JSON.parse(res.getContentText()).error.message; } catch (e) {}
      throw new Error('文字の読み取りに失敗しました（' + res.getResponseCode() + '）' + why);
    }
    const docId = JSON.parse(res.getContentText()).id;
    try {
      out.text = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=text/plain`, { headers: auth })
        .getContentText('UTF-8');
    } finally {
      DriveApp.getFileById(docId).setTrashed(true);   // 読み取り用のドキュメントは要らない
    }
  } catch (err) {
    out.ocrError = String(err.message || err);
  }
  return out;
}

// 写真から入れかけて「やめる」を押したとき。「領収書」フォルダの写真だけゴミ箱へ
function discardReceipt_(photoId) {
  if (!photoId) return;
  const file = DriveApp.getFileById(photoId);
  if (inReceiptFolder_(file)) file.setTrashed(true);
}

// 立替を保存したら、写真の名前を「日付_支払先_金額円.jpg」にそろえる
function nameReceipt_(row) {
  const m = String(row['写真'] || '').match(/[-\w]{25,}/);
  if (!m) return;
  try {
    const file = DriveApp.getFileById(m[0]);
    if (!inReceiptFolder_(file)) return;
    const d = ymd_(row['日付']);
    const day = d ? String(d[0]).slice(2) + String(d[1]).padStart(2, '0') + String(d[2]).padStart(2, '0') : '';
    const who = String(row['支払先'] || row['内容'] || '').replace(/[\\/:*?"<>|]/g, '').slice(0, 30);
    file.setName(`${day}_${who}_${Number(row['金額']) || 0}円.jpg`);
  } catch (e) {}  // 名前を変えられなくても記録は残す
}

// 最初に一度、エディタでこれを実行すると許可画面が出る（Drive と外部への接続の許可）
function setup() {
  DATA_SHEETS.forEach(sheetOf_);
  DriveApp.getRootFolder();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}
