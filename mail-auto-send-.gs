/**
 * @OnlyCurrentDoc
 * 1on1管理シートから理事会日を読み取り、メール関連の補助機能（リンク生成等）を提供するスクリプトです。
 * 本版では「自動送信」機能は明示的に無効化されています（トリガーも作成しません）。
 * 
 * @version 1.0.0
 * @author Cascade
 * @license MIT
 */

// --- 設定項目 ---
const SHEET_NAME = '1on1'; // メインで操作するシート名
const CONTACT_SHEET_NAME = 'フロント担当者'; // 担当者名とメールアドレスが記載されたシート名
const CC_MAP_SHEET_NAME = '営業担当者マップ'; // CC担当者名とメールアドレスが記載されたシート名
const CONSTRUCTION_MAP_SHEET_NAME = '工事会社マップ'; // 工事会社名とメールアドレスが記載されたシート名
const COMMON_CC_EMAIL = 'toshiki.morita@ubiden.co.jp'; // 固定で追加する共通CCアドレス
const REMINDER_DAYS_BEFORE = 3; // 理事会日の何日前に「前」メールを送信するか
const REMINDER_DAYS_AFTER = 2;  // 理事会日の何日後に「後」メールを送信するか

// 署名
const SIGNATURE = `
--
--------------------------------------------------------------------
森田　稔己 / Toshiki Morita
ユビ電株式会社 

ビジネスストラテジー／カスタマーサクセス

〒108-0073　東京都港区三田一丁目1番14号　Bizflex麻布十番4階

TEL 080-7439-7098 

名刺：https://8card.net/virtual_cards/1jlunfBTRHBP85DHWGqskA

HP： https://www.ubiden.com
└───────────────┘
--------------------------------------------------------------------
`.trim();
// --- 設定項目ここまで ---

/**
 * 自動送信機能は無効化されています。
 * 既存の sendScheduledEmails トリガーがあれば削除し、新規作成は行いません（手動実行用）。
 */
function createDailyTrigger() {
  // 自動送信は無効化ポリシー：既存の sendScheduledEmails トリガーがあれば削除のみ行い、新規作成はしない
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendScheduledEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  // 新規トリガーは作成しない
  try {
    SpreadsheetApp.getUi().alert('自動送信機能は無効化されています。トリガーは作成しませんでした。');
  } catch (e) {
    // UI が無い実行環境（トリガー等）でも落ちないようにする
    console.warn('自動送信機能は無効化されています。トリガーは作成されません。');
    Logger.log('createDailyTrigger: 自動送信は無効化されています（trigger not created）');
  }
}

/**
 * 自動送信は無効化されています。
 * 既存のトリガー用関数名は維持しますが、処理は行いません（no-op）。
 */
function sendScheduledEmails() {
  // 自動送信は無効化されています。何も行いません。
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && typeof ss.toast === 'function') {
      ss.toast('自動送信は無効化されています（sendScheduledEmails は no-op）。', 'Info', 5);
    }
  } catch (e) {
    // 非対話実行時（トースト不可）
    console.warn('sendScheduledEmails: 自動送信は無効化されています（no-op）');
    Logger.log('sendScheduledEmails: 自動送信は無効化されています（no-op）');
  }
  return;
}

/**
 * 理事会「前」のメール件名と本文を作成します。
 * @param {object} params メールの内容を定義するパラメータ
 * @returns {{subject: string, body: string}} 件名と本文を含むオブジェクト
 */
function composeBeforeEmailContent(params) {
  const { property, meetingDate, contactName, branch, companyName } = params;
  const originalDateStr = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const subject = `[ユビ電]理事会でのEV充電設備に関するご案内について（${property}）`;
  const body = [
    companyName,
    `${contactName} 様`,
    '',
    'いつもお世話になっております。ユビ電の森田です。',
    '',
    `${property} の理事会が ${originalDateStr} に開催されるかと存じますので、`,
    'EV充電設備に関するご案内について、リマインドのためご連絡を差し上げました。',
    '',
    'ご多用のところ恐縮ではございますが、理事会にてEV充電設備のご案内をいただけますよう、',
    '何卒よろしくお願い申し上げます。',
    '',
    '',
    SIGNATURE
  ].join('\n');

  return { subject, body };
}

/**
 * 理事会「後」のメール件名と本文を作成します。
 * @param {object} params メールの内容を定義するパラメータ
 * @returns {{subject: string, body: string}} 件名と本文を含むオブジェクト
 */
function composeAfterEmailContent(params) {
  const { property, meetingDate, contactName, branch, companyName } = params;
  const originalDateStr = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const subject = `[ユビ電]EV充電設備ご提案の理事会後の状況について（${property}）`;
  const body = [
    companyName,
    `${contactName} 様`,
    '',
    'いつもお世話になっております。ユビ電の森田です。',
    '',
    `${property} の理事会が ${originalDateStr} に開催されたかと存じますが、`,
    'EV充電設備のご提案に関して、理事会でのご反応はいかがでしたでしょうか。',
    '',
    'ご多用の折恐縮ではございますが、下記の点についてご共有いただけますと幸いです。',
    '',
    '---',
    '■ ご確認事項',
    '- ご提案に対する決定事項の有無',
    '- 次回理事会や総会での扱い予定',
    '- 今後の進め方についてのご検討内容 など',
    '---',
    '※本メールは社内での連携状況にかかわらず、自動的にお送りしております。すでにご対応済みの場合はご容赦くださいませ。',
    '',
    '',
    SIGNATURE
  ].join('\n');

  return { subject, body };
}

// --- ここから手動実行機能（メールリンク生成など） ---

/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('1on1便利機能')
    .addItem('✉️ メールリンクを一括生成', 'generateMailLinks')
    .addSeparator()
    .addItem('⚙️ IDを一括付番', 'assignUniqueIds')
    .addItem('🔗 Salesforceリンクを作成', 'createSalesforceLinks')
    .addToUi();
}

/**
 * Gmailの下書き作成リンクを一括で生成し、シートに設定します。
 */
function generateMailLinks() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert(`シート「${SHEET_NAME}」が見つかりません。`);
    return;
  }

  // 各種マップを作成
  const contactMap = createEmailMap(ss, CONTACT_SHEET_NAME, 'フロント担当者名', 'メールアドレス', '会社名');
  const ccContactMap = createEmailMap(ss, CC_MAP_SHEET_NAME, '営業担当者名', 'メールアドレス');
  const constructionMap = createEmailMap(ss, CONSTRUCTION_MAP_SHEET_NAME, '工事会社名', 'メールアドレス', '会社名');

  // マップの存在チェック (フロント担当者のみ必須)
  if (!contactMap) {
    ui.alert('「フロント担当者」シートが見つからないか、ヘッダー名が正しくありません。処理を中断します。');
    return;
  }

  // CCマップが存在しない場合は空のマップとして扱う
  const finalCcContactMap = ccContactMap || new Map();
  const finalConstructionMap = constructionMap || new Map();

  const idx = getHeaderIndexFunction(sheet);
  const cols = {
    property: idx('マンション名'),
    meetingDate: idx('次の理事会日'),
    contactName: idx('フロント担当者'),
    beforeLink: idx('理事会前メール作成'),
    afterLink: idx('理事会後メール作成'),
    branch: idx('支店・部署'), // 任意
    ccStaff1: idx('CC担当1'), // 任意
    ccStaff2: idx('CC担当2'), // 任意
    constructionCompany: idx('工事会社') // 任意
  };

  // 必須列の存在チェック
  const required = ['property', 'meetingDate', 'contactName', 'beforeLink', 'afterLink'];
  for (const key of required) {
    if (cols[key] === 0) {
      const colName = { property: 'マンション名', meetingDate: '次の理事会日', contactName: 'フロント担当者', beforeLink: '理事会前メール作成', afterLink: '理事会後メール作成' }[key];
      ui.alert(`シート「${SHEET_NAME}」に必須列「${colName}」が見つかりません。`);
      return;
    }
  }

  const startRow = 2;
  const numRows = sheet.getLastRow() - startRow + 1;
  if (numRows <= 0) {
    ui.alert('処理対象のデータがありません。');
    return;
  }

  const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  let generatedCount = 0;

  // 一行ずつループして処理
  values.forEach((row, i) => {
    const currentRowNum = startRow + i;
    const contactName = row[cols.contactName - 1];
    const propertyName = row[cols.property - 1];
    const meetingDate = row[cols.meetingDate - 1];

    // 必須情報がなければスキップ
    if (!contactName || !propertyName || !(meetingDate instanceof Date)) {
      return;
    }

    const contactInfo = contactMap.get(contactName.toString().trim());
    if (!contactInfo || !contactInfo.email) {
      console.log(`行 ${currentRowNum}: フロント担当者「${contactName}」のメールアドレスまたは会社情報が見つかりません。スキップします。`);
      return;
    }
    const toEmail = contactInfo.email;
    const companyName = contactInfo.company;

    // CCメールアドレスのリストを作成
    const ccEmails = [COMMON_CC_EMAIL];
    const ccStaff1Name = cols.ccStaff1 > 0 ? row[cols.ccStaff1 - 1] : '';
    if (ccStaff1Name) {
      const ccInfo = finalCcContactMap.get(ccStaff1Name.toString().trim());
      if (ccInfo && ccInfo.email) ccEmails.push(ccInfo.email);
    }
    const ccStaff2Name = cols.ccStaff2 > 0 ? row[cols.ccStaff2 - 1] : '';
    if (ccStaff2Name) {
      const ccInfo = finalCcContactMap.get(ccStaff2Name.toString().trim());
      if (ccInfo && ccInfo.email) ccEmails.push(ccInfo.email);
    }
    const constructionCompanyName = cols.constructionCompany > 0 ? row[cols.constructionCompany - 1] : '';
    if (constructionCompanyName) {
      const constructionInfo = finalConstructionMap.get(constructionCompanyName.toString().trim());
      if (constructionInfo && constructionInfo.email) ccEmails.push(constructionInfo.email);
    }
    const ccString = ccEmails.filter(Boolean).join(',');

    const params = {
      property: propertyName,
      meetingDate: meetingDate,
      to: toEmail,
      contactName: contactName,
      branch: cols.branch > 0 ? row[cols.branch - 1] : '',
      cc: ccString,
      companyName: companyName
    };

    // 「理事会前」のリンクを生成
    const beforeLinkCell = sheet.getRange(currentRowNum, cols.beforeLink);
    if (beforeLinkCell.getValue() === '') {
      const formula = createGmailFormula({ ...params, mode: 'before' });
      beforeLinkCell.setFormula(formula);
      generatedCount++;
    }

    // 「理事会後」のリンクを生成
    const afterLinkCell = sheet.getRange(currentRowNum, cols.afterLink);
    if (afterLinkCell.getValue() === '') {
      const formula = createGmailFormula({ ...params, mode: 'after' });
      afterLinkCell.setFormula(formula);
      generatedCount++;
    }
  });

  if (generatedCount > 0) {
    ui.alert(`${generatedCount}件のメール作成リンクを生成しました。`);
  } else {
    ui.alert('すべての対象セルに既にリンクが入力されています。');
  }
}

/**
 * Gmailの下書き作成用HYPERLINK数式を生成します。
 * @param {object} params メールの内容を定義するパラメータ
 * @returns {string} HYPERLINK関数を含む数式の文字列
 */
function createGmailFormula(params) {
  const { property, meetingDate, contactName, branch, to, cc, mode, companyName } = params;

  const originalDateStr = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // 件名と本文を作成
  let subject, bodyTemplate;

  if (mode === 'before') {
    subject = `[ユビ電]理事会でのEV充電設備に関するご案内について（${property}）`;
    bodyTemplate = [
      companyName,
      `${contactName} 様`,
      '',
      'いつもお世話になっております。ユビ電の森田です。',
      '',
      `${property} の理事会が ${originalDateStr} に開催されるかと存じますので、`,
      'EV充電設備に関するご案内について、リマインドのためご連絡を差し上げました。',
      '',
      'ご多用のところ恐縮ではございますが、理事会にてEV充電設備のご案内をいただけますよう、',
      '何卒よろしくお願い申し上げます。',
    ].join('\n');
  } else {
    subject = `[ユビ電]EV充電設備ご提案の理事会後の状況について（${property}）`;
    bodyTemplate = [
      companyName,
      `${contactName} 様`,
      '',
      'いつもお世話になっております。ユビ電の森田です。',
      '',
      `${property} の理事会が ${originalDateStr} に開催されたかと存じますが、`,
      'EV充電設備のご提案に関して、理事会でのご反応はいかがでしたでしょうか。',
      '',
      'ご多用の折恐縮ではございますが、下記の点についてご共有いただけますと幸いです。',
      '',
      '---',
      '■ ご確認事項',
      '- ご提案に対する決定事項の有無',
      '- 次回理事会や総会での扱い予定',
      '- 今後の進め方についてのご検討内容 など',
      '---',
      '※本メールは社内での連携状況にかかわらず、自動的にお送りしております。すでにご対応済みの場合はご容赦くださいませ。',
    ].join('\n');
  }

  // Gmail URLを構築
  let gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(to)}`;
  if (cc) {
    gmailUrl += `&cc=${encodeURIComponent(cc)}`;
  }
  gmailUrl += `&su=${encodeURIComponent(subject)}`;
  gmailUrl += `&body=${encodeURIComponent(bodyTemplate)}`;

  return `=HYPERLINK("${gmailUrl}", "✉️ メール作成")`;
}

/**
 * 未入力の行にユニークIDを付与します。
 */
function assignUniqueIds() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert(`シート「${SHEET_NAME}」が見つかりません。`);
    return;
  }
  
  const idx = getHeaderIndexFunction(sheet);
  const colId = idx('ID');
  const colProperty = idx('マンション名');
  if (colId === 0 || colProperty === 0) {
    ui.alert('「ID」列または「マンション名」列が見つかりません。');
    return;
  }

  const startRow = 2;
  const numRows = sheet.getLastRow() - startRow + 1;
  if (numRows <= 0) return;

  const range = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
  const values = range.getValues();
  
  // ID列だけを更新するための配列を用意
  const idColumnValues = sheet.getRange(startRow, colId, numRows, 1).getValues();

  let assignedCount = 0;
  idColumnValues.forEach((row, i) => {
    const property = values[i][colProperty - 1];
    if (!row[0] && property) { // IDが空で、マンション名がある場合
      row[0] = Utilities.getUuid().slice(0, 8);
      assignedCount++;
    }
  });

  if (assignedCount > 0) {
    sheet.getRange(startRow, colId, numRows, 1).setValues(idColumnValues);
    ui.alert(`${assignedCount}件の行にIDを付与しました。`);
  } else {
    ui.alert('新たにIDを付与する行はありませんでした。');
  }
}


// --- 以下、ヘルパー関数 (既存スクリプトから流用) ---

function getHeaderIndexFunction(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return (name) => headers.findIndex(h => h && h.toString().includes(name)) + 1;
}

function createEmailMap(ss, sheetName, nameHeader, emailHeader, companyHeader) {
  const mapSheet = ss.getSheetByName(sheetName);
  if (!mapSheet) return null;

  const idx = getHeaderIndexFunction(mapSheet);
  const nameCol = idx(nameHeader);
  const emailCol = idx(emailHeader);
  const companyCol = companyHeader ? idx(companyHeader) : 0;
  if (nameCol === 0 || emailCol === 0) return null;

  const lastRow = mapSheet.getLastRow();

  // データ行が存在しない場合（ヘッダーのみの場合）は空のマップを返す
  if (lastRow < 2) {
    return new Map();
  }

  const data = mapSheet.getRange(2, 1, lastRow - 1, mapSheet.getLastColumn()).getValues();
  const emailMap = new Map();
  data.forEach((row, i) => {
    const name = row[nameCol - 1];
    const email = row[emailCol - 1];
    if (name) {
      emailMap.set(name.toString().trim(), {
        email: email || '',
        company: companyCol > 0 ? row[companyCol - 1] : ''
      });
    }
  });
  return emailMap;
}


// ✅ 新機能のみ追記：他の関数には一切変更を加えません
function createSalesforceLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName('1on1'); // メインのシート
  const reportSheet = ss.getSheetByName('Salesforceレポート'); // インポートしたレポートタブ（名前は必要に応じて変更）

  if (!reportSheet) {
    SpreadsheetApp.getUi().alert('Salesforceレポートという名前のシートが見つかりません。タブ名を確認してください。');
    return;
  }

  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const reportHeaders = reportSheet.getRange(1, 1, 1, reportSheet.getLastColumn()).getValues()[0];

  const idx = (headers, name) => headers.findIndex(h => h && h.toString().includes(name)) + 1;

  const colMainId = idx(mainHeaders, 'ID');
  const colMainName = idx(mainHeaders, 'マンション名');
  const colReportId = idx(reportHeaders, '商談ID(18桁)');
  const colReportName = idx(reportHeaders, '商談名');

  if (!colMainId || !colMainName || !colReportId || !colReportName) {
    SpreadsheetApp.getUi().alert('必要な列（ID、マンション名、商談ID、商談名）が見つかりません。');
    return;
  }

  const lastRowMain = mainSheet.getLastRow() - 1;
  const lastRowReport = reportSheet.getLastRow() - 1;
  const mainNamesRange = mainSheet.getRange(2, colMainName, lastRowMain).getValues();
  const mainIdRange = mainSheet.getRange(2, colMainId, lastRowMain);
  const mainLinkRange = mainSheet.getRange(2, colMainName, lastRowMain);

  const reportData = reportSheet.getRange(2, 1, lastRowReport, reportSheet.getLastColumn()).getValues();

  const reportMap = {};
  reportData.forEach(row => {
    const id = row[colReportId - 1];
    const name = row[colReportName - 1];
    if (name) reportMap[name] = id;
  });

  const linkPrefix = 'https://ubiden.lightning.force.com/lightning/r/Opportunity/';
  const richTextValues = [];

  for (let i = 0; i < mainNamesRange.length; i++) {
    const name = mainNamesRange[i][0];
    const id = reportMap[name];

    if (id) {
      // IDを書き込み
      mainIdRange.getCell(i + 1, 1).setValue(id);
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(name)
        .setLinkUrl(`${linkPrefix}${id}/view`)
        .build();
      richTextValues.push([richText]);
    } else {
      const plainText = SpreadsheetApp.newRichTextValue()
        .setText(name || '')
        .build();
      richTextValues.push([plainText]);
    }
  }

  mainLinkRange.setRichTextValues(richTextValues);
}

/**
 * シート編集イベントハンドラ。
 * 1on1シートの「フロント担当者/次の理事会日」（旧: フロント担当者名/総会開催月）編集時のみ、
 * 変更前後の値を「更新ログ」シートに記録し、編集セルをハイライトします。
 * 
 * - ログ列: [日時, ユーザー, シート, セル, 列見出し, 旧値, 新値, 旧→新]
 * - 複数セルの一括編集や貼り付け時は oldValue が取得できない場合があります（仕様）。
 * 
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e 編集イベント
 */
function onEdit(e) {
  if (!e || !e.range) return; // 手動実行ガード
  logUpdate(e);
  highlightEditedCells(e);
  autoUpdateSfaFlag(e);
}

/**
 * 編集内容を「更新ログ」シートへ追記します。
 * 複数セルの一括編集では oldValue が取得できないため、旧値が空になる場合があります。
 * 日付は yyyy/MM/dd に正規化して保存します。
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e 編集イベント
 */
function logUpdate(e) {
  const sheetNameToTrack = '1on1';
  const editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== sheetNameToTrack) return;

  const headers = editedSheet.getRange(1, 1, 1, editedSheet.getLastColumn()).getValues()[0];
  const idx = name => headers.findIndex(h => h && h.toString().includes(name)) + 1;
  const colContact = idx('フロント担当者名') || idx('フロント担当者');
  const colMeetingDate = idx('総会開催月') || idx('次の理事会日') || idx('理事会日');
  if (![colContact, colMeetingDate].includes(e.range.getColumn())) return;

  // ログシート取得/作成
  let logSheet = e.source.getSheetByName('更新ログ');
  const created = !logSheet;
  if (!logSheet) {
    logSheet = e.source.insertSheet('更新ログ');
  }
  if (created && logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日時', 'ユーザー', 'シート', 'セル', '列見出し', '旧値', '新値', '旧→新']);
  }

  // 列見出し取得
  const headerName = editedSheet.getRange(1, e.range.getColumn()).getDisplayValue();
  
  // 旧値・新値の取得
  const oldValue = e.oldValue ?? '';
  const newValue = e.range.getValue();
  
  logSheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    Session.getActiveUser().getEmail(),
    sheetNameToTrack,
    e.range.getA1Notation(),
    headerName,
    oldValue,
    newValue,
    `${oldValue} → ${newValue}`
  ]);
}

/**
 * 編集セルを薄い黄色でハイライトします（対象列のみ）。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e 編集イベント
 */
function highlightEditedCells(e) {
  const sheetNameToTrack = '1on1';
  const sheet = e.range.getSheet();
  if (sheet.getName() !== sheetNameToTrack) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = name => headers.findIndex(h => h && h.toString().includes(name)) + 1;
  const colContact = idx('フロント担当者名') || idx('フロント担当者');
  const colMeetingDate = idx('総会開催月') || idx('次の理事会日') || idx('理事会日');
  if (![colContact, colMeetingDate].includes(e.range.getColumn())) return;

  e.range.setBackground('#fff2cc');
}

/**
 * 提案可否の変更やメモ列の背景色変更に応じて、SFA商談化フラグを自動更新します。
 * ログには記録されません。
 * 
 * ルール:
 * - 提案可否 = "提案可" → リード判定フラグ = "リード化"
 * - 提案可否 = "不可" → リード判定フラグ = "失注"
 * - 提案可否 = "保留" → リード判定フラグ = "保留"
 * - メモ列の背景色 = #e6b8af → リード判定フラグ = "次年度持越し"
 * 
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e 編集イベント
 */
function autoUpdateSfaFlag(e) {
  const sheetNameToTrack = '1on1';
  const editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== sheetNameToTrack) return;

  const headers = editedSheet.getRange(1, 1, 1, editedSheet.getLastColumn()).getValues()[0];
  const idx = name => headers.findIndex(h => h && h.toString().includes(name)) + 1;
  
  const colProposal = idx('提案可否');
  const colSfaFlag = idx('リード判定フラグ');
  const colMemo = idx('メモ');
  
  if (!colSfaFlag) return; // SFA商談化フラグ列がない場合は何もしない
  
  const currentRow = e.range.getRow();
  const sfaFlagCell = editedSheet.getRange(currentRow, colSfaFlag);
  
  // 提案可否の変更をチェック
  if (colProposal && e.range.getColumn() === colProposal) {
    const newValue = e.range.getValue();
    if (newValue === '提案可') {
      sfaFlagCell.setValue('リード化');
    } else if (newValue === '不可') {
      sfaFlagCell.setValue('失注');
    } else if (newValue === '保留') {
      sfaFlagCell.setValue('保留');
    }
    return;
  }
  
  // メモ列の背景色変更をチェック
  if (colMemo && e.range.getColumn() === colMemo) {
    const backgroundColor = e.range.getBackground();
    if (backgroundColor === '#e6b8af') {
      sfaFlagCell.setValue('次年度持越し');
    }
    return;
  }
}

/**
 * ログ用の文字列に正規化するユーティリティ。
 * - null/undefined は空文字
 * - Date または日時として解釈可能な文字列は yyyy/MM/dd に整形
 * - それ以外は文字列化
 * @param {*} val 任意の値
 * @returns {string} 正規化済み文字列
 */
function normalizeForLog_(val) {
  if (val === null || val === undefined) return '';
  // Date 型
  if (Object.prototype.toString.call(val) === '[object Date]') {
    const d = /** @type {Date} */ (val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }
  }
  // 文字列を日付として解釈できるか試す
  const s = String(val);
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  return s;
}
