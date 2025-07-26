/**
 * @OnlyCurrentDoc
 * 1on1管理シートから理事会日を読み取り、リマインドメールを自動送信するスクリプトです。
 * 毎日定時に実行されることを想定しています。
 * 
 * @version 1.7.0
 * @author Cascade
 * @license MIT
 */

// --- 設定項目 ---
const SHEET_NAME = '1on1'; // メインで操作するシート名
const CONTACT_SHEET_NAME = 'フロント担当者'; // 担当者名とメールアドレスが記載されたシート名
const CC_MAP_SHEET_NAME = '営業担当者マップ'; // CC担当者名とメールアドレスが記載されたシート名
const CONSTRUCTION_MAP_SHEET_NAME = '工事会社マップ'; // 工事会社名とメールアドレスが記載されたシート名
const COMMON_CC_EMAIL = 'sales@ubiden.com'; // 固定で追加する共通CCアドレス
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
 * 日次トリガーを設定するための関数です。手動で一度実行してください。
 */
function createDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendScheduledEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  // 新しいトリガーを設定（毎日午前9時〜10時）
  ScriptApp.newTrigger('sendScheduledEmails')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  SpreadsheetApp.getUi().alert('毎日午前9時〜10時にメールを自動送信する設定が完了しました。');
}

/**
 * スケジュールに基づいてリマインドメールを送信します。
 * この関数がトリガーによって自動実行されます。
 */
function sendScheduledEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.error(`シート「${SHEET_NAME}」が見つかりません。`);
    return;
  }

  // 各種メールアドレスのマップを作成
  const contactMap = createEmailMap(ss, CONTACT_SHEET_NAME, 'フロント担当者名', 'メールアドレス', '会社名');
  const ccContactMap = createEmailMap(ss, CC_MAP_SHEET_NAME, '営業担当者名', 'メールアドレス');
  const constructionMap = createEmailMap(ss, CONSTRUCTION_MAP_SHEET_NAME, '工事会社名', 'メールアドレス', '会社名');

  if (!contactMap || !ccContactMap || !constructionMap) {
    console.error('「フロント担当者」「営業担当者マップ」「工事会社マップ」のいずれかのシートまたはヘッダーが正しくありません。');
    return;
  }

  const idx = getHeaderIndexFunction(sheet);
  const cols = {
    property: idx('マンション名'),
    meetingDate: idx('次の理事会日'),
    contactName: idx('フロント担当者'),
    beforeSentDate: idx('理事会前メール送信日'), // ★要追加: 送信済みかを記録する列
    afterSentDate: idx('理事会後メール送信日'),   // ★要追加: 送信済みかを記録する列
    disableSend: idx('自動送信無効'),   // ★追加: 自動送信を無効にするフラグ列
    ccStaff1: idx('CC担当1'),
    ccStaff2: idx('CC担当2'),
    constructionCompany: idx('工事会社'),
    branch: idx('支店・部署'),
  };

  // 必須列の存在チェック
  const requiredCols = ['property', 'meetingDate', 'contactName', 'beforeSentDate', 'afterSentDate'];
  for (const key of requiredCols) {
    if (cols[key] === 0) {
      console.error(`必須列が見つかりません: ${key}`);
      return;
    }
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0); // 時刻をリセットして日付のみで比較

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  data.forEach((row, i) => {
    const currentRowNum = i + 2;
    const meetingDate = row[cols.meetingDate - 1];
    const propertyName = row[cols.property - 1];
    const contactName = row[cols.contactName - 1];
    const beforeSentDate = row[cols.beforeSentDate - 1];
    const afterSentDate = row[cols.afterSentDate - 1];

    // 必須情報がなければスキップ
    if (!(meetingDate instanceof Date) || !propertyName || !contactName) {
      return;
    }

    const disableSendFlag = cols.disableSend > 0 ? row[cols.disableSend - 1] : false;
    if (disableSendFlag === true) {
      return; // 自動送信無効フラグが立っていればスキップ
    }

    const contactInfo = contactMap.get(contactName.toString().trim());
    if (!contactInfo || !contactInfo.email) {
      console.log(`行 ${currentRowNum}: フロント担当者「${contactName}」のメールアドレスまたは会社情報が見つかりません。`);
      return;
    }
    const toEmail = contactInfo.email;
    const companyName = contactInfo.company;

    const commonParams = {
      property: propertyName,
      meetingDate: meetingDate,
      contactName: contactName,
      branch: cols.branch > 0 ? row[cols.branch - 1] : '',
      companyName: companyName
    };

    // CCメールアドレスのリストを作成 (共通化)
    const ccEmails = [COMMON_CC_EMAIL];
    const ccStaff1Name = cols.ccStaff1 > 0 ? row[cols.ccStaff1 - 1] : '';
    if (ccStaff1Name) {
      const ccInfo = ccContactMap.get(ccStaff1Name.toString().trim());
      if (ccInfo && ccInfo.email) ccEmails.push(ccInfo.email);
    }
    const ccStaff2Name = cols.ccStaff2 > 0 ? row[cols.ccStaff2 - 1] : '';
    if (ccStaff2Name) {
      const ccInfo = ccContactMap.get(ccStaff2Name.toString().trim());
      if (ccInfo && ccInfo.email) ccEmails.push(ccInfo.email);
    }
    const constructionCompanyName = cols.constructionCompany > 0 ? row[cols.constructionCompany - 1] : '';
    if (constructionCompanyName) {
      const constructionInfo = constructionMap.get(constructionCompanyName.toString().trim());
      if (constructionInfo && constructionInfo.email) ccEmails.push(constructionInfo.email);
    }
    const ccString = ccEmails.filter(Boolean).join(',');

    // --- 理事会「前」のメール送信チェック ---
    if (!beforeSentDate) {
      const beforeSendDate = new Date(meetingDate.getTime());
      beforeSendDate.setDate(beforeSendDate.getDate() - REMINDER_DAYS_BEFORE);
      beforeSendDate.setHours(0, 0, 0, 0);

      if (today.getTime() === beforeSendDate.getTime()) {
        const emailContent = composeBeforeEmailContent(commonParams);
        try {
          GmailApp.sendEmail(toEmail, emailContent.subject, emailContent.body, { cc: ccString });
          sheet.getRange(currentRowNum, cols.beforeSentDate).setValue(new Date());
          console.log(`行 ${currentRowNum} (${propertyName}) の理事会「前」メールを送信しました。`);
        } catch (e) {
          console.error(`行 ${currentRowNum} の理事会「前」メール送信に失敗しました: ${e.message}`);
        }
      }
    }

    // --- 理事会「後」のメール送信チェック ---
    if (!afterSentDate) {
      const afterSendDate = new Date(meetingDate.getTime());
      afterSendDate.setDate(afterSendDate.getDate() + REMINDER_DAYS_AFTER);
      afterSendDate.setHours(0, 0, 0, 0);

      if (today.getTime() === afterSendDate.getTime()) {
        const emailContent = composeAfterEmailContent(commonParams);
        try {
          GmailApp.sendEmail(toEmail, emailContent.subject, emailContent.body, { cc: ccString });
          sheet.getRange(currentRowNum, cols.afterSentDate).setValue(new Date());
          console.log(`行 ${currentRowNum} (${propertyName}) の理事会「後」メールを送信しました。`);
        } catch (e) {
          console.error(`行 ${currentRowNum} の理事会「後」メール送信に失敗しました: ${e.message}`);
        }
      }
    }
  });
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
    'いつもお世話になっております。',
    'ユビ電の営業サポートの森田です。',
    '',
    `${originalDateStr}開催予定の${property} 理事会にて、EV充電設備に関するご案内をお願いしたく、リマインドのご連絡を差し上げました。`,
    'ご多用中恐れ入りますが、よろしくお願いいたします。',
    '',
    '※理事会日程は以前の面談時に伺った内容をもとに記載しております。',
    '未定の場合は月初で仮設定しておりますので、確定日や変更等ございましたらご教示いただけますと幸いです。',
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
    'いつもお世話になっております。ユビ電の営業サポートの森田です。',
    '',
    `${property} の理事会が ${originalDateStr} に開催されたかと存じますが、`,
    'EV充電設備のご提案に関して、理事会でのご反応はいかがでしたでしょうか。',
    '',
    'ご多用の折恐縮ではございますが、下記の点についてご共有いただけますと幸いです。',
    '',
    '---',
    '■ ご確認事項',
    '- ご提案の有無',
    '- 次回理事会や総会での扱い予定',
    '- 今後の進め方についてのご検討内容 など',
    '---',
    '※本メールは社内での連携状況にかかわらずお送りしております。すでにご対応済みの場合はご容赦くださいませ。',
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
      'いつもお世話になっております。',
      'ユビ電の営業サポートの森田です。',
      '',
      `${originalDateStr}開催予定の${property} 理事会にて、EV充電設備に関するご案内をお願いしたく、リマインドのご連絡を差し上げました。`,
      'ご多用中恐れ入りますが、よろしくお願いいたします。',
      '',
      '※理事会日程は以前の面談時に伺った内容をもとに記載しております。',
      '未定の場合は月初で仮設定しておりますので、確定日や変更等ございましたらご教示いただけますと幸いです。'
    ].join('\n');
  } else {
    subject = `[ユビ電]EV充電設備ご提案の理事会後の状況について（${property}）`;
    bodyTemplate = [
      companyName,
      `${contactName} 様`,
      '',
      'いつもお世話になっております。ユビ電の営業サポートの森田です。',
      '',
      `${property} の理事会が ${originalDateStr} に開催されたかと存じますが、`,
      'EV充電設備のご提案に関して、理事会でのご反応はいかがでしたでしょうか。',
      '',
      'ご多用の折恐縮ではございますが、下記の点についてご共有いただけますと幸いです。',
      '',
      '---',
      '■ ご確認事項',
      '- ご提案の有無',
      '- 次回理事会や総会での扱い予定',
      '- 今後の進め方についてのご検討内容 など',
      '---',
      '※本メールは社内での連携状況にかかわらずお送りしております。すでにご対応済みの場合はご容赦くださいませ。',
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
