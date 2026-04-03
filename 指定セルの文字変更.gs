function applyTextsToMultipleSpreadsheets() {
  // =========================
  // 設定
  // =========================

  // 対象スプレッドシートURLを複数入れる
  const spreadsheetUrls = [
        "https://docs.google.com/spreadsheets/d/1P8nXyk9ruBCDOzmyCGKUst3KMJ6HUOjpL1eqMEssJlM/edit?usp=drivesdk",
        "https://docs.google.com/spreadsheets/d/1miqIx-lrrYnSMl-sNDcTPgrzENgAo0QVKilhWSxsbU0/edit?usp=drivesdk",
        "https://docs.google.com/spreadsheets/d/1giyUwJkvJIzucyQQsvvUdmst9Ons99jcKLL9MsvTE2s/edit?usp=drivesdk",
  ];

  // 対象シート名を複数入れる
  const targetSheetNames = [
    '月次管理シートテンプレ_各人',
    '26年2月'
  ];

  // =========================
  // 変更したいセルと入れる文字（複数可）
  // キー = セル
  // 値   = 入れたい文字
  // =========================
  const textMap = {
    'E3': '予定稼働時間',
  };

  // =========================
  // 実行
  // =========================
  spreadsheetUrls.forEach(url => {
    const ss = SpreadsheetApp.openById(getSpreadsheetIdFromUrl_(url));

    targetSheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`シートが見つかりません: ${sheetName} / ${url}`);
        return;
      }

      applyTextsToSheet_(sheet, textMap);

      Logger.log(`完了: ${ss.getName()} / ${sheetName}`);
    });
  });
}


/**
 * 指定シートに文字を入れる
 */
function applyTextsToSheet_(sheet, textMap) {
  Object.keys(textMap).forEach(a1Notation => {
    sheet.getRange(a1Notation).setValue(textMap[a1Notation]);
  });
}


/**
 * スプレッドシートURLからIDを抜き出す
 */
function getSpreadsheetIdFromUrl_(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error(`スプレッドシートURLが不正です: ${url}`);
  }
  return match[1];
}
