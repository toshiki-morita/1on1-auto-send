/**
 * @OnlyCurrentDoc
 * 1on1管理シートから理事会日を読み取り、リマインドメールを自動送信するスクリプトです。
 * 毎日定時に実行されることを想定しています。
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
  const contactMap = createEmailMap(ss, CONTACT_SHEET_NAME, 'フロント担当者名', 'メールアドレス');
  const ccContactMap = createEmailMap(ss, CC_MAP_SHEET_NAME, '営業担当者名', 'メールアドレス');
  const constructionMap = createEmailMap(ss, CONSTRUCTION_MAP_SHEET_NAME, '工事会社名', 'メールアドレス');

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

    const commonParams = {
      property: propertyName,
      meetingDate: meetingDate,
      contactName: contactName,
      branch: cols.branch > 0 ? row[cols.branch - 1] : '',
      companyName: ss.getName()
    };

    const toEmail = contactMap.get(contactName.toString().trim());
    if (!toEmail) {
      console.log(`行 ${currentRowNum}: フロント担当者「${contactName}」のメールアドレスが見つかりません。`);
      return;
    }

    // CCメールアドレスのリストを作成 (共通化)
    const ccEmails = [COMMON_CC_EMAIL];
    if (cols.ccStaff1 > 0) {
      const ccStaff1Name = row[cols.ccStaff1 - 1];
      if (ccStaff1Name) ccEmails.push(ccContactMap.get(ccStaff1Name.toString().trim()));
    }
    if (cols.ccStaff2 > 0) {
      const ccStaff2Name = row[cols.ccStaff2 - 1];
      if (ccStaff2Name) ccEmails.push(ccContactMap.get(ccStaff2Name.toString().trim()));
    }
    if (cols.constructionCompany > 0) {
      const companyName = row[cols.constructionCompany - 1];
      if (companyName) ccEmails.push(constructionMap.get(companyName.toString().trim()));
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
    `${companyName}   ${branch}`.trim(),
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
    `${companyName}   ${branch}`.trim(),
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

  // 各種マップを作成 (シートやヘッダーがなくてもエラーにしない)
  const contactMap = createEmailMap(ss, CONTACT_SHEET_NAME, 'フロント担当者名', 'メールアドレス') || new Map();
  const ccContactMap = createEmailMap(ss, CC_MAP_SHEET_NAME, '営業担当者名', 'メールアドレス') || new Map();
  const constructionMap = createEmailMap(ss, CONSTRUCTION_MAP_SHEET_NAME, '工事会社名', 'メールアドレス') || new Map();

  // フロント担当者マップは必須とする
  if (contactMap.size === 0 && (ss.getSheetByName(CONTACT_SHEET_NAME) === null || ss.getSheetByName(CONTACT_SHEET_NAME).getLastRow() < 2)) {
     ui.alert('「フロント担当者」シートにデータがありません。処理を中断します。');
     return;
  }

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

    const toEmail = contactMap.get(contactName.toString().trim());
    if (!toEmail) {
      console.log(`行 ${currentRowNum}: フロント担当者「${contactName}」のメールアドレスが見つかりません。スキップします。`);
      return;
    }

    // CCメールアドレスのリストを作成
    const ccEmails = [COMMON_CC_EMAIL];
    if (cols.ccStaff1 > 0) {
      const ccStaff1Name = row[cols.ccStaff1 - 1];
      if (ccStaff1Name) ccEmails.push(ccContactMap.get(ccStaff1Name.toString().trim()));
    }
    if (cols.ccStaff2 > 0) {
      const ccStaff2Name = row[cols.ccStaff2 - 1];
      if (ccStaff2Name) ccEmails.push(ccContactMap.get(ccStaff2Name.toString().trim()));
    }
    if (cols.constructionCompany > 0) {
      const companyName = row[cols.constructionCompany - 1];
      if (companyName) ccEmails.push(constructionMap.get(companyName.toString().trim()));
    }
    const ccString = ccEmails.filter(Boolean).join(',');

    const params = {
      property: propertyName,
      meetingDate: meetingDate,
      to: toEmail,
      contactName: contactName,
      branch: cols.branch > 0 ? row[cols.branch - 1] : '',
      cc: ccString,
      companyName: ss.getName()
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
      `${companyName}   ${branch}`.trim(),
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
      `${companyName}   ${branch}`.trim(),
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

function createEmailMap(ss, sheetName, nameHeader, emailHeader) {
  const mapSheet = ss.getSheetByName(sheetName);
  if (!mapSheet) return null;

  const idx = getHeaderIndexFunction(mapSheet);
  const colName = idx(nameHeader);
  const colEmail = idx(emailHeader);
  if (colName === 0 || colEmail === 0) return null;

  const lastRow = mapSheet.getLastRow();

  // データ行が存在しない場合（ヘッダーのみの場合）は空のマップを返す
  if (lastRow < 2) {
    return new Map();
  }

  const data = mapSheet.getRange(2, 1, lastRow - 1, mapSheet.getLastColumn()).getValues();
  const emailMap = new Map();
  data.forEach(row => {
    const name = row[colName - 1];
    const email = row[colEmail - 1];
    if (name && email) {
      emailMap.set(name.toString().trim(), email.toString().trim());
    }
  });
  return emailMap;
}
