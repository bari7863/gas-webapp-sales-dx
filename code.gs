function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('営業管理グラフ & ランキング');
}

// Webアプリ用：当日グラフ・月間グラフのデータをまとめて返す
function getAllChartData() {
  var ss = SpreadsheetApp.getActive();

  return {
    daily:  parseChartSheet_(ss, '当日グラフ'),
    monthly: parseChartSheet_(ss, '月間グラフ')
  };
}

function updateGraphNumbersOnly() {
  var ss = SpreadsheetApp.getActive();

  SpreadsheetApp.flush();

  return {
    daily:  parseChartSheet_(ss, '当日グラフ'),
    monthly: parseChartSheet_(ss, '月間グラフ')
  };
}

// 「当日グラフ」「月間グラフ」の構造を解析して JSON に変換
function parseChartSheet_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null; // ヘッダーのみ

  // A〜E列を読み込む（氏名, 項目, ラベル, 目標, 実数）
  var values = sh.getRange(1, 1, lastRow, 5).getValues();

  var grouped = [];

  // 1行目はヘッダーなので 2行目（index=1）から読む
  for (var i = 1; i < values.length; i++) {
    var r = values[i];

    // 全部空ならスキップ
    var isEmpty = r.every(function (c) {
      return c === '' || c === null;
    });
    if (isEmpty) continue;

    grouped.push({
      name:   r[0],                                    // 氏名
      metric: r[1],                                    // 項目名（コール数 / 商談数 / 受注数）
      label:  r[2] || (r[0] + ' / ' + r[1]),           // ラベルが空なら「氏名 / 項目」
      goal:   Number(r[3]) || 0,                       // 目標
      actual: Number(r[4]) || 0                        // 実数
    });
  }

  if (grouped.length === 0) return null;

  // 以前の構造との互換性のため work/order/sales も空配列で返す
  return {
    grouped: grouped,
    work:    [],
    order:   [],
    sales:   []
  };
}
