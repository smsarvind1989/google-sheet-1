# google-sheet-1
active cells layout in sheet script
// [] = saari sheets, [3] = sirf Sr 2 (Amardeep verma) wali sheet
const ONLY_ROWS = [];

function formatLinkedSheets() {
  const mainSS = SpreadsheetApp.getActiveSpreadsheet();
  const emp = mainSS.getSheetByName('Employees');
  const lastRow = emp.getLastRow();
  if (lastRow < 2) return;

  // Column H = "Employee Sheet" (Open sheet links), row 2 se shuru
  const rich = emp.getRange(2, 8, lastRow - 1, 1).getRichTextValues();
  const forms = emp.getRange(2, 8, lastRow - 1, 1).getFormulas();

  const items = [];
  for (let i = 0; i < rich.length; i++) {
    const row = i + 2;
    if (ONLY_ROWS.length > 0 && ONLY_ROWS.indexOf(row) === -1) continue;

    let url = rich[i][0].getLinkUrl();
    if (!url && forms[i][0]) {
      const m = forms[i][0].match(/HYPERLINK\("([^"]+)"/i);
      if (m) url = m[1];
    }
    if (url) items.push({ row: row, url: url });
  }

  items.forEach(item => {
    try {
      const idMatch = item.url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!idMatch) throw new Error('ID nahi mili: ' + item.url);
      const ss = SpreadsheetApp.openById(idMatch[1]);
      ss.getSheets().forEach(sheet => formatSheet(sheet));
      Logger.log('Done - row ' + item.row + ' : ' + ss.getName());
    } catch (e) {
      Logger.log('Error - row ' + item.row + ' : ' + e);
    }
  });
}

function formatSheet(sheet) {
  const range = sheet.getDataRange();   // A1 se last data cell tak
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;
  const lastCol = range.getLastColumn();

  // 1. Borders all sides - poore data area par (khaali cells bhi)
  range.setBorder(true, true, true, true, true, true);

  // 2. Wrap text
  range.setWrap(true);

  // 3. Header row (row 1) - bold + light background
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#d9e2f3');

  // 4. Top row freeze
  sheet.setFrozenRows(1);

  // 5. Autofit columns (row height wrap ke saath apne aap adjust hoti hai)
  sheet.autoResizeColumns(1, lastCol);
}
