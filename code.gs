function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Seperate')
  .addItem('insert basis of increment', 'insertHorizontalBordersAfterMajorIncrements')
  .addItem('insert basis of grouping', 'insertHorizontalBordersAfterGrouping')
  .addToUi();
}


function insertHorizontalBordersAfterMajorIncrements() {
  const spreadsheet = SpreadsheetApp.openById('1OMlwBBhX6TVraZNOmlYUpV931jjdLiD_HdyqHSRnTE8');
  const sheet = spreadsheet.getActiveSheet();
  const lastRow = sheet.getLastRow();

  let previousMajorIncrement = '';
  let rowsInserted = 0;

  for (let i = 2; i <= lastRow + rowsInserted; i++) {
    const cell = sheet.getRange(i, 1);
    const cellValue = cell.getValue().toString();
    const majorIncrement = cellValue.split('.')[0];

    if (majorIncrement !== previousMajorIncrement) {
      if (i > 2) {
        // Apply a horizontal border to the row above the major increment.
        sheet.getRange(i - 1, 1, 1, sheet.getMaxColumns()).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
      }

      previousMajorIncrement = majorIncrement;
    }
  }
}

function insertHorizontalBordersAfterGrouping() {
  const spreadsheet = SpreadsheetApp.openById('1OMlwBBhX6TVraZNOmlYUpV931jjdLiD_HdyqHSRnTE8');
  const sheet = spreadsheet.getActiveSheet();
  const lastRow = sheet.getLastRow();

  let previousFirstLetter = '';
  let rowsInserted = 0;

  for (let i = 2; i <= lastRow + rowsInserted; i++) {
    const firstLetter = sheet.getRange(i, 2).getValue().charAt(0);

    if (firstLetter !== previousFirstLetter) {
      if (i > 2) {
        // Apply a horizontal border to the row above the new first letter in column B.
        sheet.getRange(i - 1, 1, 1, sheet.getMaxColumns()).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
      }

      previousFirstLetter = firstLetter;
    }
  }
}
