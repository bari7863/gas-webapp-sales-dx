function updateDynamicSumFormulasForRange() {
  // =========================
  // 設定
  // =========================

  // 対象スプレッドシートURL
  const spreadsheetUrls = [
    'https://docs.google.com/spreadsheets/d/1qeceAOApZM0KVvGOGKjOlkub5KbYOabh6OOBBvyfHUc/edit?gid=293816191#gid=293816191',
  ];

  // 対象シート名
  const targetSheetNames = [
    '26年2月',
  ];

  // 対象列：H～AL
  const startCol = 8;   // H
  const endCol   = 38;  // AL

  // 関数を入れ替える行範囲：H4:AL18
  const formulaStartRow = 4;
  const formulaEndRow   = 18;

  // 参照開始行
  // 4行目 → 23行
  // 5行目 → 24行
  // ...
  // 18行目 → 37行
  const sourceStartRowBase = 23;

  // 何行おきか
  const step = 19;

  // ここを好きな行数に変更する
  // 例: 306, 363, 500
  const targetMaxRow = 306;

  // C21 用
  const countaTargetCell = 'C21';
  const countaSourceColumn = 'E';
  const countaStartRow = 22;

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

      updateFormulaRange_(
        sheet,
        startCol,
        endCol,
        formulaStartRow,
        formulaEndRow,
        sourceStartRowBase,
        step,
        targetMaxRow
      );

      updateCountaFormula_(
        sheet,
        countaTargetCell,
        countaSourceColumn,
        countaStartRow,
        targetMaxRow,
        step
      );

      Logger.log(`完了: ${ss.getName()} / ${sheetName}`);
    });
  });
}


/**
 * H4:AL18 のような範囲全体を更新
 */
function updateFormulaRange_(
  sheet,
  startCol,
  endCol,
  formulaStartRow,
  formulaEndRow,
  sourceStartRowBase,
  step,
  targetMaxRow
) {
  const rowCount = formulaEndRow - formulaStartRow + 1;
  const colCount = endCol - startCol + 1;

  const formulas2D = [];

  for (let formulaRow = formulaStartRow; formulaRow <= formulaEndRow; formulaRow++) {
    const currentRowFormulas = [];

    // 4行目→23, 5行目→24 ... 18行目→37
    const startSourceRow = sourceStartRowBase + (formulaRow - formulaStartRow);

    for (let col = startCol; col <= endCol; col++) {
      const colLetter = columnToLetter_(col);
      const formula = buildDynamicSumFormula_(colLetter, startSourceRow, targetMaxRow, step);
      currentRowFormulas.push(formula);
    }

    formulas2D.push(currentRowFormulas);
  }

  sheet.getRange(formulaStartRow, startCol, rowCount, colCount).setFormulas(formulas2D);
}


/**
 * C21 の COUNTA 関数を更新
 * 例:
 * =COUNTA(E22,E41,E60,...)
 */
function updateCountaFormula_(sheet, targetCellA1, sourceColumn, startRow, maxRow, step) {
  const refs = [];

  for (let row = startRow; row <= maxRow; row += step) {
    refs.push(`${sourceColumn}${row}`);
  }

  const formula = refs.length > 0 ? `=COUNTA(${refs.join(',')})` : '';
  sheet.getRange(targetCellA1).setFormula(formula);
}


/**
 * 例:
 * =SUM(H23,H42,H61,...)
 */
function buildDynamicSumFormula_(colLetter, startRow, maxRow, step) {
  const refs = [];

  for (let row = startRow; row <= maxRow; row += step) {
    refs.push(`${colLetter}${row}`);
  }

  if (refs.length === 0) {
    return '';
  }

  return `=SUM(${refs.join(',')})`;
}


/**
 * URLからスプレッドシートIDを取得
 */
function getSpreadsheetIdFromUrl_(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error(`スプレッドシートURLが不正です: ${url}`);
  }
  return match[1];
}


/**
 * 列番号 → 列記号
 * 例: 8 → H
 */
function columnToLetter_(column) {
  let temp = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    temp = String.fromCharCode(65 + remainder) + temp;
    column = Math.floor((column - 1) / 26);
  }
  return temp;
}
