/**
 * 設定項目
 * 機密情報はGASの「プロジェクトの設定 > スクリプトプロパティ」から取得する。
 */
const props = PropertiesService.getScriptProperties();

const CONFIG = {
  // スクリプトプロパティから取得した値をセット
  TEMPLATE_ID: props.getProperty('TEMPLATE_ID'), 
  ROOT_FOLDER_ID: props.getProperty('ROOT_FOLDER_ID'),
  ADMIN_EMAIL: props.getProperty('ADMIN_EMAIL'),
  SHEET_NAME: '研修管理シート'
};

/**
 * 列定義（データ定義シートに基づく：A列=0スタート）
 */
const COLS = {
  USER_ID: 0,         // A: 受講者ID
  LAST_NAME: 1,       // B: 姓
  FIRST_NAME: 2,      // C: 名
  COURSE_NAME: 3,     // D: コース名
  STATUS: 4,          // E: ステータス
  COMPLETION_DATE: 5, // F: 修了日
  FILE_ID: 6,         // G: 修了証ID
  EMAIL: 7,           // H: メールアドレス
  IS_APPROVED: 8      // I: 発行承認 (Boolean)
};

/**
 * 1. トリガー関数: スプレッドシートの編集時に実行
 * 【重要】この関数は「インストーラブルトリガー（編集時）」として設定してください。
 */
function onUpdateStatus(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  if (row < 2) return;

//選択された範囲を一列目から一行分だけ取得する
  const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  const values = range.getValues()[0];

  const status = values[COLS.STATUS];
  const isApproved = values[COLS.IS_APPROVED];
  const currentFileId = values[COLS.FILE_ID];

  // 起動条件: ステータスが「受講完了」かつ 承認フラグが TRUE かつ IDが存在しない
  if (status === '受講完了' && isApproved === true && !currentFileId) {
    try {
      createCertificate(sheet, row, values);
    } catch (err) {
      console.error(err);
      // 設計書に基づき管理者へ通知（ここではログとメッセージボックス）
      if (typeof Browser !== 'undefined') {
        Browser.msgBox(`証明書作成エラー(行:${row}): ${err.message}`);
      }
    }
  }
}

/**
 * 修了証生成処理
 */
function createCertificate(sheet, row, data) {
  const userId = data[COLS.USER_ID];
  const lastName = data[COLS.LAST_NAME];
  const firstName = data[COLS.FIRST_NAME];
  const courseName = data[COLS.COURSE_NAME];

  if (!userId || !lastName || !firstName || !courseName) {
    throw new Error('必須項目（ID, 姓名, コース名）が不足しています。');
  }

  let completionDate = data[COLS.COMPLETION_DATE];
  const today = new Date();
  if (!completionDate) {
    completionDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy年MM月dd日');
    sheet.getRange(row, COLS.COMPLETION_DATE + 1).setValue(completionDate);
  } else {
    completionDate = Utilities.formatDate(new Date(completionDate), Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  }

  const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
  const fileName = `修了証_${courseName}_${lastName}${firstName}_${userId}`;

  const docFile = templateFile.makeCopy(fileName, rootFolder);
  const docId = docFile.getId();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();//Googleドキュメントの中身を操作できる一句

  body.replaceText('{lastName}', lastName);
  body.replaceText('{firstName}', firstName);
  body.replaceText('{courseName}', courseName);
  body.replaceText('{completionDate}', completionDate);
  doc.saveAndClose();

  const pdfBlob = docFile.getAs(MimeType.PDF);//拡張子を指定する一句
  rootFolder.createFile(pdfBlob).setName(fileName + '.pdf');

  sheet.getRange(row, COLS.FILE_ID + 1).setValue(docId);
}

/**
 * 2. バッチ処理: メール一括送信
 * 【重要】この関数は「時間主導型トリガー」として設定してください。
 */
/**
 * 2. バッチ処理: メール一括送信（氏名・ID不備も検知する強化版）
 */
function sendBatchEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  //指定した名前のシートを探す
  if (!sheet) {
    const names = ss.getSheets().map(s => s.getName()).join(', ');
    //join:シート名の配列をカンマ区切りの文字列に変換（バラバラのデータを文章化）
    throw new Error(`シート「${CONFIG.SHEET_NAME}」が見つかりません。存在するシート: [${names}]`);
  }

  const lastRow = sheet.getLastRow();//読み込む行数を最適化
  if (lastRow < 2) return;

  // 2行目から最後まで、データの範囲を特定
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  //特定した範囲のデータを全部取り出す
  const values = range.getValues();
  
  //集計用カウンターの準備（結果報告メールで使用）
  let sentCount = 0;
  let errorCount = 0;
  let errorMessages = [];

  //全員に同じ処理をする
  values.forEach((row, index) => {
    const realRowIndex = index + 2;
    const status = row[COLS.STATUS];
    const isApproved = row[COLS.IS_APPROVED];
    const fileId = row[COLS.FILE_ID];
    const email = row[COLS.EMAIL];
    const lastName = row[COLS.LAST_NAME] || "（氏名未入力）"; // レポート用に代替文字を設定
    const firstName = row[COLS.FIRST_NAME];
    const courseName = row[COLS.COURSE_NAME];
    const userId = row[COLS.USER_ID];

    // 送信対象の基本判定：ステータスが「受講完了」かつ「承認」にチェックがある場合
    if (status === '受講完了' && isApproved === true) {
      
      try {
        // --- 必須項目の一括バリデーション ---
        let missingFields = [];//空のリストがあればここに追加
        if (!userId) missingFields.push("受講者ID");
        if (!lastName || lastName === "（氏名未入力）") missingFields.push("姓");
        if (!firstName) missingFields.push("名");
        if (!courseName) missingFields.push("コース名");
        if (!email || email.toString().trim() === "") missingFields.push("メールアドレス");
        
        if (missingFields.length > 0) {
          throw new Error(`必須項目不足: [${missingFields.join(", ")}]`);
        }

        // 修了証ファイルがまだ生成されていない場合のチェック
        if (!fileId) {
          throw new Error("修了証PDFが生成されていません（自動生成トリガーが失敗している可能性があります）");
        }

        // --- 送信処理 ---
        const docFile = DriveApp.getFileById(fileId);
        const pdfBlob = docFile.getAs(MimeType.PDF).setName(`修了証_${courseName}.pdf`);

        const subject = `【修了証送付】${courseName} の受講完了のお知らせ`;
        const body = `${lastName} 様\n\nお疲れ様です。\n\n以下の研修コースの受講が完了いたしました。\n修了証を添付いたしますので、ご確認ください。\n\nコース名: ${courseName}\n\n以上、よろしくお願いいたします。`;

        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body,
          attachments: [pdfBlob]
        });

        // 送信成功時の更新
        sheet.getRange(realRowIndex, COLS.STATUS + 1).setValue('修了証送信済み');
        sheet.getRange(realRowIndex, COLS.IS_APPROVED + 1).setValue(false);
        sentCount++;

      } catch (e) {
        // すべての不備をここでキャッチしてエラーレポートに載せる
        console.error(`Row ${realRowIndex} Error: ${e.message}`);
        errorCount++;
        errorMessages.push(`行 ${realRowIndex} (${lastName}): ${e.message}`);
      }
    }
  });

  // 管理者に結果報告メール
  if (sentCount > 0 || errorCount > 0) {
    let reportBody = `本日の送信処理が完了しました。\n\n`;
    reportBody += `・送信成功: ${sentCount} 件\n`;
    reportBody += `・エラー: ${errorCount} 件\n\n`;
    
    if (errorMessages.length > 0) {
      reportBody += `■エラー詳細:\n` + errorMessages.join('\n');
    }

    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: '【システム通知】修了証一括送信レポート',
      body: reportBody
    });
  }
}
