function autoFillAndClean() {

  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();

  const numRows = values.length;
  const numCols = values[0].length;

  // --- 1. 1行目で最も右にデータがある列を探す ---
  let lastCol = -1;
  for (let c = numCols - 1; c >= 0; c--) {
    if (values[0][c] !== "") {
      lastCol = c;
      break;
    }
  }
  if (lastCol < 1) return;

  const prevCol = lastCol - 1;

  // --- 2. 自動補完処理 ---
  for (let r = 1; r < numRows; r++) {

    const arrowText = values[r][lastCol];

    if (typeof arrowText === "string" && arrowText !== "") {

      // 直前列が空なら
      if (values[r][prevCol] === "") {

        // 左へ遡って値を探す
        for (let c = prevCol - 1; c >= 0; c--) {
          if (values[r][c] !== "") {

            const cell = sheet.getRange(r + 1, prevCol + 1);
            cell.setValue(values[r][c]);
            cell.setFontColor("#888888"); // グレー文字

            break;
          }
        }
      }
    }
  }

  // --- 3. 最新2列より左のグレー文字を削除 ---
  for (let r = 1; r < numRows; r++) {
    for (let c = 0; c < prevCol - 1; c++) {

      const cell = sheet.getRange(r + 1, c + 1);
      const fontColor = cell.getFontColor();

      // 自動生成グレーのみ削除
      if (fontColor === "#888888") {
        cell.setValue("");
      }
    }
  }
}
