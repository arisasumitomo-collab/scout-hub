/**
 * ビズリーチ スカウト管理システム - Google Apps Script バックエンド（JSONP対応版）
 *
 * 【設置方法】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 > Apps Script を開く
 * 3. このコードをすべて貼り付けて保存
 * 4. デプロイ > 新しいデプロイ > ウェブアプリ
 *    - 「次のユーザーとして実行」: 自分
 *    - 「アクセスできるユーザー」: 全員
 * 5. デプロイして発行されたURLをアプリの「設定」タブに入力
 */

const SHEET_NAMES = {
  JOBS:       'jobs',
  ACTUALS:    'actuals',
  SCOUT_DATA: 'scout',
};

// ─── メインエントリーポイント（JSONP対応）────────────────
function doGet(e) {
  const callback = e.parameter.callback || '';
  const postdata = e.parameter.postdata;
  let result;

  try {
    if (postdata) {
      // POSTデータをGETパラメータで受け取る（JSONP対応）
      const body = JSON.parse(postdata);
      const action = body.action;
      if      (action === 'saveJob')       result = saveJob(body.data);
      else if (action === 'saveActual')    result = saveActual(body.data);
      else if (action === 'saveScoutData') result = saveScoutData(body.data);
      else if (action === 'deleteJob')     result = deleteJob(body.jobId);
      else                                 result = { error: 'unknown action: ' + action };
    } else {
      const action = e.parameter.action;
      if      (action === 'getAll')  result = getAllData();
      else if (action === 'getJobs') result = getJobs();
      else                           result = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── シート初期化 ─────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

// ─── 求人マスタ ────────────────────────────────────────────
const JOB_HEADERS = [
  'job_id', 'position_name', 'job_type', 'start_date',
  'weekly_target', 'monthly_target', 'is_active', 'created_at'
];

function getJobs() {
  const sheet = getOrCreateSheet(SHEET_NAMES.JOBS, JOB_HEADERS);
  return sheetToObjects(sheet);
}

function saveJob(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.JOBS, JOB_HEADERS);
  const rows = sheetToObjects(sheet);
  const existing = rows.findIndex(r => r.job_id === data.job_id);
  const row = [
    data.job_id, data.position_name, data.job_type, data.start_date,
    data.weekly_target || 0, data.monthly_target || 0,
    data.is_active !== false ? 'TRUE' : 'FALSE',
    data.created_at || new Date().toISOString()
  ];
  if (existing >= 0) {
    sheet.getRange(existing + 2, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true };
}

function deleteJob(jobId) {
  const sheet = getOrCreateSheet(SHEET_NAMES.JOBS, JOB_HEADERS);
  const rows = sheetToObjects(sheet);
  const idx = rows.findIndex(r => r.job_id === jobId);
  if (idx >= 0) sheet.deleteRow(idx + 2);
  return { success: true };
}

// ─── 実績データ ────────────────────────────────────────────
const ACTUAL_HEADERS = [
  'actual_id', 'job_id', 'period_type', 'period_label',
  'sent_count', 'created_at'
];

function saveActual(data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.ACTUALS, ACTUAL_HEADERS);
  const rows = sheetToObjects(sheet);
  const existing = rows.findIndex(
    r => r.job_id === data.job_id &&
         r.period_type === data.period_type &&
         r.period_label === data.period_label
  );
  const row = [
    data.actual_id || Utilities.getUuid(),
    data.job_id, data.period_type, data.period_label,
    data.sent_count, new Date().toISOString()
  ];
  if (existing >= 0) {
    sheet.getRange(existing + 2, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true };
}

// ─── CSVスカウトデータ ─────────────────────────────────────
const SCOUT_HEADERS = [
  'job_id', 'position_name', 'job_type',
  'platinum_sent', 'platinum_read', 'platinum_read_rate',
  'platinum_reply', 'platinum_reply_rate',
  'total_entry', 'doc_pass', 'interview1', 'interview2',
  'final_interview', 'offer', 'hire',
  'imported_at'
];

function saveScoutData(dataArray) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SCOUT_DATA, SCOUT_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  dataArray.forEach(d => {
    sheet.appendRow([
      d.job_id, d.position_name, d.job_type,
      d.platinum_sent, d.platinum_read, d.platinum_read_rate,
      d.platinum_reply, d.platinum_reply_rate,
      d.total_entry, d.doc_pass, d.interview1, d.interview2,
      d.final_interview, d.offer, d.hire,
      new Date().toISOString()
    ]);
  });
  return { success: true, count: dataArray.length };
}

// ─── 全データ一括取得 ──────────────────────────────────────
function getAllData() {
  return {
    jobs:      getJobs(),
    actuals:   getActuals(),
    scoutData: getScoutData(),
  };
}

function getActuals() {
  const sheet = getOrCreateSheet(SHEET_NAMES.ACTUALS, ACTUAL_HEADERS);
  return sheetToObjects(sheet);
}

function getScoutData() {
  const sheet = getOrCreateSheet(SHEET_NAMES.SCOUT_DATA, SCOUT_HEADERS);
  return sheetToObjects(sheet);
}

// ─── ユーティリティ ────────────────────────────────────────
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
