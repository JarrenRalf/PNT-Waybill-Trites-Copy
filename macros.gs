/**
* Simple function to add a menu option to the spreadsheet "Export", for saving a PDF of the spreadsheet directly to Google Drive.
* The exported file will be named: SheetName and saved in the same folder as the spreadsheet.
* To change the filename, just set pdfName inside generatePdf() to something else.
* Running this, sends the currently open sheet, as a PDF attachment
*/
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('Sort Box #s on Packing Slip').addItem("Sort by Box Numbers", "sortPackingSlipByBoxNumbers").addToUi()

  //SpreadsheetApp.getActiveSpreadsheet().addMenu('Export', [{name:"Save PDF", functionName:"generatePdf"}]) 
}

function onEdit(e)
{
  var spreadsheet = SpreadsheetApp.getActive();
  var labelSheet = spreadsheet.getSheetByName('Labels');
  var printSheet = spreadsheet.getSheetByName('Print Label with Piece Count')
  var wb = spreadsheet.getRange('WAYBILL!I8').getValue() + 1;
  var ui = SpreadsheetApp.getUi();
  
  var colStart = e.range.columnStart;
  var active = e.source.getActiveSheet();
  var name = active.getName();
  var response;
  
  const NUM_ROWS_PRINT_PAGE = 35;
  const NUM_LABELS_PER_PAGE = 6;
  
  var maxNumLabels = printSheet.getMaxRows()/NUM_ROWS_PRINT_PAGE*NUM_LABELS_PER_PAGE;
  
  // Monitor the 'Number of Pieces' input on the LABEL page and react appropriately based on the value and data type of the entry
  if(e.range.getA1Notation() === 'D16')
  {
    // If input is Not a Number (NaN) or blank or contains spaces
    if (isNaN(e.range.getValue()) || e.range.isBlank() || e.range.getValue().toString().includes(" ")) 
    {
      response = ui.alert('Invalid Input!', 'You must enter a number into this cell.', ui.ButtonSet.OK);
      
      // Reset the piece count to 1
      if (response == ui.Button.OK)
        labelSheet.getRange('D16').setValue(1);
    }
    else // Otherwise the input is a number
    {
      if (e.range.getValue() <= 0) // Negative number
      {
        response = ui.alert('Invalid Input!', 'You must enter a positive number into this cell.', ui.ButtonSet.OK);
      
        // Reset the piece count to 1
        if (response == ui.Button.OK)
          labelSheet.getRange('D16').setValue(1);
      }
      else if (e.range.getValue() > maxNumLabels) // Greater than the number of labels
      {
        response = ui.alert('Too many labels!', 'You are attempting to make more labels than the current template allows for.' +
                                                ' Are you sure you would like to insert more rows?.', ui.ButtonSet.YES_NO);
        // Reset the piece count to 1
        if (response == ui.Button.NO)
          labelSheet.getRange('D16').setValue(1);
        else
          addLabelPages(maxNumLabels, e.range.getValue())
      }
    }
  }
  
  if ( name == "WAYBILL" && colStart == 16 )
  {
     // Change to your "From" sheet and Column reference
     var value = e.value; 
     
     if ( value == "saveAddress")
     {
       spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Consignee'), true);
       spreadsheet.getRange('A1').activate();
       spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
       spreadsheet.getCurrentCell().offset(1, 0).activate();
       spreadsheet.getRange('WAYBILL!J7:J10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, true);
       spreadsheet.getCurrentCell().offset(0, 4).activate();
       spreadsheet.getRange('WAYBILL!B14').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
       spreadsheet.getRange('A:F').activate().sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
       spreadsheet.getRange('A2').activate();
       spreadsheet.getRange('WAYBILL!P6').clear({contentsOnly: true, skipFilteredRows: true});
       spreadsheet.getRange('WAYBILL!B14').activate();
     }
  }
  else if ( e.range.getA1Notation() === 'J6')
  {   
    spreadsheet.getRange('WAYBILL!I8').setValue(wb)
    
    if(!(spreadsheet.getRange('J7') .getFormula().charAt(0) == '=' &&
         spreadsheet.getRange('J8') .getFormula().charAt(0) == '=' &&
         spreadsheet.getRange('J9') .getFormula().charAt(0) == '=' &&
         spreadsheet.getRange('J10').getFormula().charAt(0) == '='))
    {
      active.getRange('J7') .setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select A where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('J8') .setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select B where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('J9') .setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select C where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('J10').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select D where G like \'%"&$J$6&"%\'" ,))');
      active.getRange('B14').setFormula('=if($J$6="","x",QUERY(CONSIGNEE,"select E where G like \'%"&$J$6&"%\'" ,))');
    }
  }
  else if ( e.range.getA1Notation() === 'J8')
    spreadsheet.getRange('WAYBILL!I8').setValue(wb);
}

function getColumnWidthsAndRowHeights()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  var totalWidth = 0, totalHeight = 0;

  for (var j = 1; j <= sheet.getMaxColumns(); j++)
  {
    totalWidth += sheet.getColumnWidth(j);
    Logger.log('Col ' + j + ': ' + sheet.getColumnWidth(j))
  }
    
  for (var i = 1; i <= sheet.getMaxRows(); i++)
  {
    totalHeight += sheet.getRowHeight(i);
    Logger.log('Row ' + i + ': ' + sheet.getRowHeight(i))
  }
  
  Logger.log('totalWidth: ' + totalWidth)
  Logger.log('totalHeight: ' + totalHeight)
}

function saveLabelAddress()
{
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Consignee'), true);
    spreadsheet.getRange('A1').activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('Labels!B9:B12').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, true);
    spreadsheet.getRange('A:F').activate().sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
    spreadsheet.getRange('A2').activate();
    spreadsheet.getRange('Labels!B9').activate();
}

function resetPackingS()
{
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Reset piece counts
  spreadsheet.getRange('B19').setFormula('=WAYBILL!$B18');
  spreadsheet.getRange('B20').setFormula('=WAYBILL!$B20');
  spreadsheet.getRange('B21').setFormula('=WAYBILL!$B22');
  spreadsheet.getRange('B22').setFormula('=WAYBILL!$B24');
  spreadsheet.getRange('B23').setFormula('=WAYBILL!$B26');
  spreadsheet.getRange('B24').setFormula('=WAYBILL!$B28');
  
  // Reset Descriptions
  spreadsheet.getRange('D19').setFormula('=WAYBILL!$D18');
  spreadsheet.getRange('D20').setFormula('=WAYBILL!$D20');
  spreadsheet.getRange('D21').setFormula('=WAYBILL!$D22');
  spreadsheet.getRange('D22').setFormula('=WAYBILL!$D24');
  spreadsheet.getRange('D23').setFormula('=WAYBILL!$D26');
  spreadsheet.getRange('D24').setFormula('=WAYBILL!$D28');
  
  spreadsheet.getRange('B19').activate();
}

function resetLabel()
{
  SpreadsheetApp.getActiveSheet().getRange(13, 2, 5).setFormulas([['=WAYBILL!$J7'], ['=WAYBILL!$J8'], ['=WAYBILL!$J9'], ['=WAYBILL!$J10'], ['IF(AND(D20,D21),IF(AND(ISBLANK(WAYBILL!K14),ISBLANK(WAYBILL!F14)),"",IF(ISBLANK(WAYBILL!K14),"ORD# "&WAYBILL!F14,IF(ISBLANK(WAYBILL!F14),"PO# "&WAYBILL!K14,"PO# "&WAYBILL!K14&" ORD# "&WAYBILL!F14))),IF(D20,IF(ISBLANK(WAYBILL!K14),"","PO# "&WAYBILL!K14),IF(D21,IF(ISBLANK(WAYBILL!F14),"","ORD# "&WAYBILL!F14),"")))']])
}

function printPage()
{
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Print Label').activate(), true);
}

function UntitledMacro()
{
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B19').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!B18');
  spreadsheet.getRange('D19:P19').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!D18');
  spreadsheet.getRange('B20').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!B20');
  spreadsheet.getRange('D19:P19').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('WAYBILL'), true);
  spreadsheet.getRange('D20:I21').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Consignee'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Packing Slip'), true);
  spreadsheet.getRange('D20:P20').activate();
  spreadsheet.getCurrentCell().setFormula('=WAYBILL!D22');
}

/**
* This function sets multiple lables to be printed with a line of text that has the piece count on them. 
*
* @author Jarren
*/
function multiPrintPage()
{
  var spreadsheet = SpreadsheetApp.getActive();
  var labelSheet = spreadsheet.getSheetByName('Labels');
  var printSheet = spreadsheet.getSheetByName('Print Label with Piece Count')
  var  pieceName = labelSheet.getRange('D12').getValue();
  var pieceCount = labelSheet.getRange('D16').getValue();
  var poNumber = labelSheet.getRange('B17').getValue();
  var rangesToClear = printSheet.getDataRange();
  var string = "", richTextValue, startOffset1, endOffset1, startOffset2, endOffset2;
  
  const   START_ROW =  3;
  const      H_LINE =  5; // Horizontal Line location which separates the Shipper and Consignee addresses
  const  LEFT_LABEL =  2;
  const RIGHT_LABEL =  4;
  const  LABEL_JUMP = 11; // The vertical translation of a label on the same piece of paper
  const   PAGE_JUMP =  2; // The vertical translation of the last row of labels to the first row on the next page
  const  NUM_LABELS_PER_PAGE =   6;
  const  NUM_ROWS_PRINT_PAGE = 35;
  
  var maxNumLabels = printSheet.getMaxRows()/NUM_ROWS_PRINT_PAGE*NUM_LABELS_PER_PAGE;
  
  rangesToClear.clearContent(); // Clear all information on the sheet
  
  var bold1 = SpreadsheetApp.newTextStyle().setFontSize(12).setBold(true).build();
  var bold2 = SpreadsheetApp.newTextStyle().setFontSize(15).setBold(true).build();
  var normal = SpreadsheetApp.newTextStyle().setFontSize(12).setBold(false).build();
  
  for (var i = 0; i < maxNumLabels; i++)
  {
    if (i % 2 == 0) // If even index (Left Label)
    {
      if (i < pieceCount) // Set the label
      {
        string = poNumber + "     " + pieceName + " #  " + (i + 1).toString() + "  of  " + pieceCount; // Set the text
        startOffset1 = string.length - pieceCount.toString().length - 6 - (i + 1).toString().length;
        endOffset1 = startOffset1 + (i + 1).toString().length + 1;
        startOffset2 = string.length - pieceCount.toString().length;
        endOffset2 = string.length;
        richTextValue = SpreadsheetApp.newRichTextValue().setText(string)
          .setTextStyle(normal).setTextStyle(0, (poNumber.length === 0) ? 5 : poNumber.length, bold1)
          .setTextStyle(startOffset1, endOffset1, bold2)
          .setTextStyle(startOffset2, endOffset2, bold2)
          .build();

        setLabel(START_ROW +       i/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), LEFT_LABEL, richTextValue);
      }
      else // Otherwise clear the the black background fill of the horizontal lines
        printSheet.getRange(H_LINE + i/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), LEFT_LABEL).clearFormat();
    }
    else // If odd index (Right Label)
    {
      if (i < pieceCount) // Set the label
      {
        string = poNumber + "     " + pieceName + " #  " + (i + 1).toString() + "  of  " + pieceCount; // Set the text
        startOffset1 = string.length - pieceCount.toString().length - 6 - (i + 1).toString().length;
        endOffset1 = startOffset1 + (i + 1).toString().length + 1;
        startOffset2 = string.length - pieceCount.toString().length;
        endOffset2 = string.length;
        richTextValue = SpreadsheetApp.newRichTextValue().setText(string)
          .setTextStyle(normal).setTextStyle(0, (poNumber.length === 0) ? 5 : poNumber.length, bold1)
          .setTextStyle(startOffset1, endOffset1, bold2)
          .setTextStyle(startOffset2, endOffset2, bold2)
          .build();

        setLabel(START_ROW + (i - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL, richTextValue);
      }
      else // Otherwise clear the the black background fill of the horizontal lines
        printSheet.getRange(H_LINE + (i - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL).clearFormat();
    }
  }
  
  // Take the user to the 'Print Label with Piece Count' sheet
  spreadsheet.setActiveSheet(printSheet.activate(), true);
}

/**
* This function sets multiple lables to be printed with a line of text that has the piece count on them. 
*
* @author Jarren
*/
function multiPrintPage_MultiOrder()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const printSheet = spreadsheet.getSheetByName('Print Multi Order Labels');
  const pieceName = labelSheet.getRange('D4').getValue();
  const poNumber = labelSheet.getRange('B14').getValue();
  const  bold1 = SpreadsheetApp.newTextStyle().setFontSize(12).setBold(true).build();
  const  bold2 = SpreadsheetApp.newTextStyle().setFontSize(15).setBold(true).build();
  const normal = SpreadsheetApp.newTextStyle().setFontSize(12).setBold(false).build();
  const   START_ROW =  3;
  const      H_LINE =  5; // Horizontal Line location which separates the Shipper and Consignee addresses
  const  LEFT_LABEL =  2;
  const RIGHT_LABEL =  4;
  const  LABEL_JUMP = 11; // The vertical translation of a label on the same piece of paper
  const   PAGE_JUMP =  2; // The vertical translation of the last row of labels to the first row on the next page
  const  NUM_LABELS_PER_PAGE =  6;
  const  NUM_ROWS_PRINT_PAGE = 35;
  const maxNumLabels = printSheet.getMaxRows()/NUM_ROWS_PRINT_PAGE*NUM_LABELS_PER_PAGE;
  const poOrdAndPieceCount = spreadsheet.getSheetByName('Multi PO Labels').getSheetValues(8, 4, 7, 4)
    .filter(val => val[0] !== '' && val[1] !== '' && val[3] !== '').map(val => [val[0], val[1], val[3]]);
  var string = "", richTextValue, startOffset1, endOffset1, startOffset2, endOffset2;

  printSheet.getDataRange().clearContent(); // Clear all information on the sheet
  
  var i = 0;

  for (var j = 0; j < poOrdAndPieceCount.length; j++)
  {
    for (var k = 1; k <= poAndPieceCount[j][2]; k++)
    {
      string = poNumber + "     " + pieceName + " #  " + k.toString() + "  of  " + poAndPieceCount[j][1]; // Set the text
      startOffset1 = string.length - poAndPieceCount[j][1].toString().length - 6 - (i + 1).toString().length;
      endOffset1 = startOffset1 + (i + 1).toString().length + 1;
      startOffset2 = string.length - poAndPieceCount[j][1].toString().length;
      endOffset2 = string.length;
      richTextValue = SpreadsheetApp.newRichTextValue().setText(string)
        .setTextStyle(normal).setTextStyle(0, (poAndPieceCount[j][0].length === 0) ? 5 : poAndPieceCount[j][0].length + 4, bold1)
        .setTextStyle(startOffset1, endOffset1, bold2)
        .setTextStyle(startOffset2, endOffset2, bold2)
        .build();

      if (i % 2 == 0) // If even index (Left Label)
        setLabel_QCL(START_ROW + i/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), LEFT_LABEL, richTextValue, printSheet);
      else
        setLabel_QCL(START_ROW + (i - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL, richTextValue, printSheet);

      i++;
    }
  }

  for (var h = i; h < maxNumLabels; h++)
    if (h % 2 == 0)
      printSheet.getRange(H_LINE + h/2*LABEL_JUMP + PAGE_JUMP*Math.floor(h/NUM_LABELS_PER_PAGE), LEFT_LABEL).clearFormat();
    else
      printSheet.getRange(H_LINE + (h - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(h/NUM_LABELS_PER_PAGE), RIGHT_LABEL).clearFormat();
  
  printSheet.getRange(3, 2).activate(); // Take the user to the Print QCL Labels sheet
}

/**
* This function sets multiple lables to be printed with a line of text that has the piece count on them. 
*
* @author Jarren
*/
function multiPrintPage_QCL()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const printSheet = spreadsheet.getSheetByName('Print QCL Labels');
  const  bold1 = SpreadsheetApp.newTextStyle().setFontSize(12).setBold(true).build();
  const  bold2 = SpreadsheetApp.newTextStyle().setFontSize(15).setBold(true).build();
  const normal = SpreadsheetApp.newTextStyle().setFontSize(12).setBold(false).build();
  const   START_ROW =  3;
  const      H_LINE =  5; // Horizontal Line location which separates the Shipper and Consignee addresses
  const  LEFT_LABEL =  2;
  const RIGHT_LABEL =  4;
  const  LABEL_JUMP = 11; // The vertical translation of a label on the same piece of paper
  const   PAGE_JUMP =  2; // The vertical translation of the last row of labels to the first row on the next page
  const  NUM_LABELS_PER_PAGE =  6;
  const  NUM_ROWS_PRINT_PAGE = 35;
  const maxNumLabels = printSheet.getMaxRows()/NUM_ROWS_PRINT_PAGE*NUM_LABELS_PER_PAGE;
  const poAndPieceCount = spreadsheet.getSheetByName('QCL Labels').getSheetValues(8, 4, 7, 3).filter(val => val[0] !== '' && val[2] !== '').map(val => [val[0], val[2]]);
  var string = "", richTextValue, startOffset1, endOffset1, startOffset2, endOffset2;

  printSheet.getDataRange().clearContent(); // Clear all information on the sheet
  
  var i = 0;

  for (var j = 0; j < poAndPieceCount.length; j++)
  {
    for (var k = 1; k <= poAndPieceCount[j][1]; k++)
    {
      string = "PO# " + poAndPieceCount[j][0] + "                                  Box #  " + k.toString() + "  of  " + poAndPieceCount[j][1]; // Set the text
      startOffset1 = string.length - poAndPieceCount[j][1].toString().length - 6 - (i + 1).toString().length;
      endOffset1 = startOffset1 + (i + 1).toString().length + 1;
      startOffset2 = string.length - poAndPieceCount[j][1].toString().length;
      endOffset2 = string.length;
      richTextValue = SpreadsheetApp.newRichTextValue().setText(string)
        .setTextStyle(normal).setTextStyle(0, (poAndPieceCount[j][0].length === 0) ? 5 : poAndPieceCount[j][0].length + 4, bold1)
        .setTextStyle(startOffset1, endOffset1, bold2)
        .setTextStyle(startOffset2, endOffset2, bold2)
        .build();

      if (i % 2 == 0) // If even index (Left Label)
        setLabel_QCL(START_ROW + i/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), LEFT_LABEL, richTextValue, printSheet);
      else
        setLabel_QCL(START_ROW + (i - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL, richTextValue, printSheet);

      i++;
    }
  }

  for (var h = i; h < maxNumLabels; h++)
    if (h % 2 == 0)
      printSheet.getRange(H_LINE + h/2*LABEL_JUMP + PAGE_JUMP*Math.floor(h/NUM_LABELS_PER_PAGE), LEFT_LABEL).clearFormat();
    else
      printSheet.getRange(H_LINE + (h - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(h/NUM_LABELS_PER_PAGE), RIGHT_LABEL).clearFormat();
  
  printSheet.getRange(3, 2).activate(); // Take the user to the Print QCL Labels sheet
}

/**
* This function prints the label with a piece count.
*
* @param row    The row that the label starts at
* @param col    The column of the label
* @param string The string representing the piece count
* @author Jarren Ralf
*/
function setLabel(row, col, string)
{
  var printSheet = SpreadsheetApp.getActive().getSheetByName('Print Label with Piece Count');
  
  printSheet.getRange(row    , col).setFormula('=WAYBILL!B6'); // PNT Logo
  printSheet.getRange(row + 2, col).setBackground("black");    // Horizontal black line
  printSheet.getRange(row + 4, col).setFontWeight("normal");
  printSheet.getRange(row + 4, col).setVerticalAlignment("top")
  printSheet.getRange(row + 4, col).setValue(" ship to:");
  printSheet.getRange(row + 5, col).setFormula('=CONCATENATE("    ",Labels!$B13)');
  printSheet.getRange(row + 6, col).setFormula('=CONCATENATE("    ",Labels!$B14)');
  printSheet.getRange(row + 7, col).setFormula('=CONCATENATE("    ",Labels!$B15)');
  printSheet.getRange(row + 8, col).setFormula('=CONCATENATE("    ",Labels!$B16)');
  printSheet.getRange(row + 9, col).setHorizontalAlignment("right");
  printSheet.getRange(row + 9, col).setRichTextValue(string); // Set the piece count text
}

/**
* This function prints the label with a piece count for Queen Charlotte Lodge.
*
* @param {Number}   row     : The row that the label starts at
* @param {Number}   col     : The column of the label
* @param {String}  string   : The string representing the piece count and PO number
* @param {Sheet} printSheet : The sheet that the labels will get printed on
* @author Jarren Ralf
*/
function setLabel_QCL(row, col, string, printSheet)
{
  printSheet.getRange(row, col).setFormula("=WAYBILL!B6") // PNT Logo
    .offset(2, 0, 7).setBackgrounds([ ["black"],
                                      ["white"],
                                      ["white"],
                                      ["white"],
                                      ["white"],
                                      ["white"],
                                      ["white"]])
      .setVerticalAlignments([["bottom"],
                              ["bottom"],
                              ["top"],
                              ["bottom"],
                              ["bottom"],
                              ["bottom"],
                              ["bottom"]])
      .setFontWeights([ ["bold"],
                        ["bold"],
                        ["normal"],
                        ["bold"],
                        ["bold"],
                        ["bold"],
                        ["bold"]])
      .setValues([[""],
                  [""],
                  [" ship to:"],
                  ["    Queen Charlotte Lodge"],
                  ["    Richmond, BC  V7B 1C3"],
                  ["    Tel. (604) 420-7197"],
                  ["    Attn: John Sedo "]])
    .offset(7, 0, 1, 1).setHorizontalAlignment("right").setRichTextValue(string); // Set the piece count text
}

/**
* This function inserts and sets the row heights for additional pages of labels, the number of which is chosen by the user.
*
* @param maxNumLabels The maximum number of labels currently on the 'Print Label with Piece Count' page
* @param numLabels    The number of labels the user wants printed
* @author Jarren
*/
function addLabelPages(currentNumLabels, numLabels)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName('Print Label with Piece Count');
  var rowIndex, rowHeights, destinationRange;
  
  var numLabelsNeeded = getNumLabels(numLabels);
  
  const           FIRST_ROW =   1;
  const   NUM_ROWS_PER_PAGE =  35;
  const NUM_LABELS_PER_PAGE =   6;
  const     LAST_ROW_HEIGHT =  50;
  const    ARBITRARY_COLUMN =   2;
  
  var currentNumPages = currentNumLabels/NUM_LABELS_PER_PAGE;
  var additionalPages = (numLabelsNeeded - currentNumLabels)/NUM_LABELS_PER_PAGE;
  
  // Set the row height of the last row of each page for printing purposes
  sheet.setRowHeight(NUM_ROWS_PER_PAGE, LAST_ROW_HEIGHT);  
  
  var range = sheet.getRange(FIRST_ROW, ARBITRARY_COLUMN, NUM_ROWS_PER_PAGE); // The range of the first label page
  
  rowHeights = getRowHeights(range); // The heights of the rows on the first label page
  
  for (var j = 0; j < additionalPages; j++)
  {
    rowIndex = NUM_ROWS_PER_PAGE*(currentNumPages + j);
    sheet.insertRowsAfter(rowIndex, NUM_ROWS_PER_PAGE);
    destinationRange = sheet.getRange(rowIndex + 1, ARBITRARY_COLUMN, NUM_ROWS_PER_PAGE);
    setRowHeights(rowHeights, destinationRange);
  }
}

/**
* This funtion will return the number of labels.
*
* @param The input is some numerical value
* @return This function rounds the inputted value up to the nearest multiple of 6, and returns it's value
*/
function getNumLabels(x)
{
    return Math.ceil(x/6)*6;
}

/**
 * SET THE ROW HEIGHTS OF A SELECTED RANGE OF A DESTINATION SHEET
 * @param {Array.<number>} rHeights - row heights from getRowHeights(rng);
 * @param {object} destRange - destionation range of copied data.
 */
function setRowHeights(rHeights,destRange){
 
 var rngRowStart = destRange.getRow();
 var rngRowHeight = destRange.getHeight() + rngRowStart;
 
 var destSheet = destRange.getSheet();
 Logger.log(destSheet.getName());
 var count = 0;
 for( var i = rngRowStart; i < rngRowHeight; i++){
   destSheet.setRowHeight(i,rHeights[count]);
   
   count+=1;
 }
}

/**
 * This function rearranges the box numbers of the Packing Slip with Box Numbers sheets.
 * 
 * @author ChaptGPT
 */
function sortPackingSlipByBoxNumbers()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const numItemsPerPage = 20;
  const startRow = 18;
  const startCol = 2;
  const numCols = 12;

  // --- STEP 1: Determine number of pages ---
  const pageIndicator = spreadsheet.getSheetByName("Packing Slip with Box Numbers 1").getSheetValues(39, 5, 1, 1)[0][0];

  let totalPages = 1;

  if (pageIndicator && typeof pageIndicator === "string")
  {
    const match = pageIndicator.match(/Page\s+\d+\s+of\s+(\d+)/i);

    if (match)
      totalPages = parseInt(match[1], 10);
  }

  // --- STEP 2: Pull all data into one array ---
  let allRows = [];

  for (let page = 1; page <= totalPages; page++)
  {
    const range = spreadsheet.getSheetByName(`Packing Slip with Box Numbers ${page}`).getRange(startRow, startCol, numItemsPerPage, numCols);

    // Force text BEFORE reading
    range.setNumberFormat('@');
    SpreadsheetApp.flush()

    const data = range.getDisplayValues(); // ← THIS IS THE FIX

    data.forEach(row => allRows.push(row));
  }

  // --- STEP 3: Separate rows ---
  let validRows = [];
  let blankRows = [];

  allRows.forEach(row => {

    const qty = row[0];
    const desc = row[2];

    const isBlankRow = row.every(cell => cell === "" || cell === null);

    if (isBlankRow) {
      blankRows.push(row);
      return;
    }

    const qtyIsBlank = qty === "" || qty === null;
    const qtyIsZero = Number(qty) === 0;
    const descIsBlank = desc === "" || desc === null;

    // ❌ Remove invalid rows
    if (qtyIsBlank || qtyIsZero || descIsBlank) {
      return;
    }

    // ✅ Clean description
    row[2] = desc.toString().replace(/""+/g, '"');

    validRows.push(row);
  });

 function parseBox(value) {
  if (!value) {
    return { start: Infinity, end: Infinity, type: 3 };
  }

  const str = value.toString().trim();

  // Normalize spacing
  const clean = str.replace(/\s+/g, '');

  // Range: "2-4"
  if (clean.includes('-')) {
    const [start, end] = clean.split('-').map(x => parseInt(x, 10));
    return {
      start: start || Infinity,
      end: end || start || Infinity,
      type: 2 // range (lowest priority)
    };
  }

  // Comma: "3,4"
  if (clean.includes(',')) {
    const nums = clean.split(',').map(x => parseInt(x, 10)).filter(n => !isNaN(n));
    return {
      start: Math.min(...nums),
      end: Math.max(...nums),
      type: 1 // comma (middle)
    };
  }

  // Single: "3"
  const num = parseInt(clean, 10);
  return {
    start: isNaN(num) ? Infinity : num,
    end: isNaN(num) ? Infinity : num,
    type: 0 // single (highest priority)
  };
}

  // --- STEP 5: Sort ---
  validRows.sort((a, b) => {

    const A = parseBox(a[numCols - 1]);
    const B = parseBox(b[numCols - 1]);

    // 1. Start box
    if (A.start !== B.start)
      return A.start - B.start;

    // 2. End box (critical for overlaps)
    if (A.end !== B.end)
      return A.end - B.end;

    // 3. Type priority: single < comma < range
    if (A.type !== B.type)
      return A.type - B.type;

    // 4. Description
    return (a[2] || "").toString().toLowerCase().localeCompare((b[2] || "").toString().toLowerCase());
  });

  // --- STEP 6: Combine back ---
  const finalRows = [...validRows, ...blankRows];

  // Pad if needed
  const totalSlots = totalPages * numItemsPerPage;
  while (finalRows.length < totalSlots)
    finalRows.push(new Array(numCols).fill(""));

  // --- STEP 7: Write back to sheets ---
  for (let i = 1, index = 0; i <= totalPages; i++) {

    spreadsheet.getSheetByName(`Packing Slip with Box Numbers ${i}`).getRange(startRow, startCol, numItemsPerPage, numCols)
      .setNumberFormat('@').setValues(finalRows.slice(index, index + numItemsPerPage));

    index += numItemsPerPage;
  }
}

/**
 * GET THE ROW HEIGHTS OF A SELECTED RANGE OF A SOURCE SHEET
 * @param {object} range - Selected source range
 * @returns {Array.<number>}  Array of row heights for each row
 */
function getRowHeights(range) {
 
 var rngRowStart = range.getRow();
 var rngRowHeight = range.getHeight() + rngRowStart;
 
 var rowHeights = []
 var rangeSheet = range.getSheet()
 
 for( var i = rngRowStart; i < rngRowHeight; i++){
   var rowHeight = rangeSheet.getRowHeight(i);
   rowHeights.push(rowHeight);
 }
 return rowHeights;
}

/**
* This function generates a pdf of the current page and saves it into the same folder of the spreadsheet on the google drive.
*/
function generatePdf()
{
  // Get active spreadsheet.
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Get active sheet.
  var sheets = spreadsheet.getSheets();
  var sheetName = spreadsheet.getActiveSheet().getName();
  var sourceSheet = spreadsheet.getSheetByName(sheetName);
  
  // Set the output filename as SheetName.
  var pdfName = spreadsheet.getRange('G2').getValue();
  //var pdfName = sheetName;

  // Get folder containing spreadsheet to save pdf in.
  var parents = DriveApp.getFileById(spreadsheet.getId()).getParents();
  if (parents.hasNext())
    var folder = parents.next();
  else
    folder = DriveApp.getRootFolder();
  
  // Copy whole spreadsheet.
  var destSpreadsheet = SpreadsheetApp.open(DriveApp.getFileById(spreadsheet.getId()).makeCopy("tmp_convert_to_pdf", folder))

  // Delete redundant sheets.
  var sheets = destSpreadsheet.getSheets();
  for (i = 0; i < sheets.length; i++)
  {
    if (sheets[i].getSheetName() != sheetName)
      destSpreadsheet.deleteSheet(sheets[i]);
  }
  
  var destSheet = destSpreadsheet.getSheets()[0];

  // Repace cell values with text (to avoid broken references).
  var sourceRange = sourceSheet.getRange(1,1,sourceSheet.getMaxRows(),sourceSheet.getMaxColumns());
  var sourcevalues = sourceRange.getValues();
  var destRange = destSheet.getRange(1, 1, destSheet.getMaxRows(), destSheet.getMaxColumns());
  destRange.setValues(sourcevalues);

  // Save to pdf.
  var theBlob = destSpreadsheet.getBlob().getAs('application/pdf').setName(pdfName);
  var newFile = folder.createFile(theBlob);

  // Delete the temporary sheet.
  DriveApp.getFileById(destSpreadsheet.getId()).setTrashed(true);
}