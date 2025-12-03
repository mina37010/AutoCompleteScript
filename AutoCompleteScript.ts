function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const values = usedRange.getValues();
    const numRows = values.length;
    const numCols = values[0].length;

    // 最も右にデータが入っている列を参照
    let lastCol = -1;
    for (let c = numCols - 1; c >= 0; c--) {
        if (values[0][c] !== "") {
            lastCol = c;
            break;
        }
    }
    if (lastCol < 1) return;

    const prevCol = lastCol - 1;

    // 各行をチェックして 名前 を処理 
    for (let r = 1; r < numRows; r++) {

        let arrowText = values[r][lastCol];
        if (typeof arrowText === "string" && arrowText !="") {

            // 直前列が空欄なら自動補完
            if (values[r][prevCol] === "") {
                // 左へ遡って黒文字を捜索
                for (let c = prevCol - 1; c >= 0; c--) {
                    if (values[r][c] !== "") {

                        // 値をセット
                        let targetCell = sheet.getCell(r, prevCol);
                        targetCell.setValue(values[r][c]);

                        // テキストをグレー
                        targetCell.getFormat().getFont().setColor("#888888");

                        break;
                    }
                }
            }
        }
    }

    // 最新2列より左にある “自動生成のグレー文字” を消す
    for (let r = 1; r < numRows; r++) {
        for (let c = 0; c < prevCol - 1; c++) {

            let cell = sheet.getCell(r, c);
            let fontColor = cell.getFormat().getFont().getColor();

            // グレー文字(#888888) のみ削除
            if (fontColor === "#888888") {
                cell.setValue("");
            }
        }
    }

}
