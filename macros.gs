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
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B9') .setFormula('=WAYBILL!$J7');
  spreadsheet.getRange('B10').setFormula('=WAYBILL!$J8');
  spreadsheet.getRange('B11').setFormula('=WAYBILL!$J9');
  spreadsheet.getRange('B12').setFormula('=WAYBILL!$J10');
  spreadsheet.getRange('B13').setFormula('=CONCATENATE("",WAYBILL!$K14)');
  spreadsheet.getRange('B9').activate();
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

    if (i % 2 == 0) // If even index (Left Label)
    {
      if (i < pieceCount) // Set the label
        setLabel(START_ROW +       i/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), LEFT_LABEL, richTextValue);
      else // Otherwise clear the the black background fill of the horizontal lines
        printSheet.getRange(H_LINE + i/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), LEFT_LABEL).clearFormat();
    }
    else // If odd index (Right Label)
    {
      if (i < pieceCount) // Set the label
        setLabel(START_ROW + (i - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL, richTextValue);
      else // Otherwise clear the the black background fill of the horizontal lines
        printSheet.getRange(H_LINE + (i - 1)/2*LABEL_JUMP + PAGE_JUMP*Math.floor(i/NUM_LABELS_PER_PAGE), RIGHT_LABEL).clearFormat();
    }
  }
  
  // Take the user to the 'Print Label with Piece Count' sheet
  spreadsheet.setActiveSheet(printSheet.activate(), true);
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
  const     LAST_ROW_HEIGHT = 100;
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
* Simple function to add a menu option to the spreadsheet "Export", for saving a PDF of the spreadsheet directly to Google Drive.
* The exported file will be named: SheetName and saved in the same folder as the spreadsheet.
* To change the filename, just set pdfName inside generatePdf() to something else.
* Running this, sends the currently open sheet, as a PDF attachment
*/
function onOpen()
{
  var submenu = [{name:"Save PDF", functionName:"generatePdf"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Export', submenu);  
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