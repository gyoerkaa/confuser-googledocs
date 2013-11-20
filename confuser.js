function onOpen() {
  var activeSheet = SpreadsheetApp.getActive();
  var items = [
    {name: 'Display Matrix', functionName: 'displayTexCode'},
    {name: 'Export Matrix',  functionName: 'exportTexCode'},
  ];
  activeSheet.addMenu('Confuser', items);
}


function displayTexCode() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var texCode = generateTexCode(activeSpreadsheet);
    
  var htmlOutput = HtmlService.createHtmlOutput()
    .setSandboxMode(HtmlService.SandboxMode.NATIVE)
    .setTitle('Latex Code')
    .setWidth(900)
    .setHeight(450);
  htmlOutput.append('<div id="texCodeLineNumbers">');
  htmlOutput.append('</div>');  
  htmlOutput.append('<div id="texCodeArea" style="white-space: pre;">');
  for (var i = 0; i < texCode.length; ++i) {
      htmlOutput.appendUntrusted(texCode[i])
      htmlOutput.append('<br>');
  }

  htmlOutput.append('</div>');
  activeSpreadsheet.show(htmlOutput);
} 
    
    
function exportTexCode() {  
  var fileName = Browser.inputBox('Save tex file as: ');
  
  if (fileName.length <= 0) {
    Browser.msgBox("Error: Please enter a file name.");
    return;
  }
   
  // Append extension if not already present
  if ( (fileName.length <= 4) || (fileName.substr(fileName.length-4,4) != '.tex'))
    fileName += ".tex";
    
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var texCode = generateTexCode(activeSpreadsheet);
  DocsList.createFile(fileName, texCode.join('\\n'));  
}
 

/**
 * Get the texcode for the confusion matrix
 *
 * @param {spreadsheet} spreadsheet to get the data from
 * @return {Array} Texcode, each array element represents a single line of code.
 *                  No newlines are appended.
 */
function generateTexCode(spreadsheet) {
  var data = getCMatCells(spreadsheet.getActiveSheet());
  // check if there is actually something in the data array
  if ((data.length <= 0) || (data[0].length <= 0)) {
    spreadsheet.toast('Not enough data (3x3 cells or more required)');
  }
  var texCode = new Array();
  texCode.push('\\begin{table}[H]%');
  texCode.push('    \\centering'); 
  
  line = '    \\begin{tabular}{|cc|';
  for (c = 0; c < data[0].length-1; ++c)
    line += 'c|';
  line += '}%';
  texCode.push(line); 
  
  texCode.push('    \\cline{1-' + (data[0].length+1) + '}%');  
  texCode.push('    & & \\multicolumn{' + 
               (data[0].length-1) + 
               '}{c|}{' + 
               'Predicted' +
               '} \\\\');  
  texCode.push('    \\cline{3-' + (data[0].length+1) + '}%');  
  
  line = '    & ';
  for (c = 1; c < data[0].length; ++c)
    line += ' & ' + data[0][c];
  line += '\\\\';
  texCode.push(line); 
  
  texCode.push('    \\cline{1-' + (data[0].length+1) + '}%');  
  line = '    \\multicolumn{1}{|c}{\\multirow{' +
         data[0].length +
         '}{*}{\\begin{sideways}' +
         'Observed' +
         '\\end{sideways}}} & \\multicolumn{1}{|c|}{' +
         data[1][0] +
          '}';
  for (c = 1; c < data[0].length; ++c)
    line += ' & ' + data[1][c];
  line += '\\\\';
  texCode.push(line);
  
  for (r = 2; r < data.length; ++r) {
    line = '    \\cline{2-' + data[0].length+1 + '}%';
    texCode.push(line);

    line = '    \\multicolumn{1}{|c}{} & \\multicolumn{1}{|c|}{' +
           data[r][0] +
           '}';
    for (c = 1; c < data[0].length; ++c)
      line += ' & ' + data[r][c];
    line += '\\\\';
    texCode.push(line);
  }
    
  texCode.push('    \\cline{1-' + data[0].length + '}%'); 
  texCode.push('    \\end{tabular}%');
  texCode.push('    \\caption[Short Caption]{Long Caption}%');
  texCode.push('    \\label{tab:Label}');
  texCode.push('\\end{table}');  
  
  return texCode;
}


/**
 * Get the cells in the spreadsheet to use for the
 * confustion matrix. The data will come from the 
 * selected/active cells unless the selected range
 * is to small. If the selected range is smaller 
 * than 3, all cells will be used.
 * The returned Array will be empty, if there are 
 * not enough entries/cells.
 *
 * @param {spreadsheet} spreadsheet to get the data from
 * @return {DataTable} DataTable of elemts to use for the confusion matrix
 */
function getCMatCells(sheet) {
  var range = sheet.getActiveRange();
  if ((range.getHeight() < 3) || (range.getWidth() < 3)) {
    // cannot create matrix from this, try to get all cells
    range = sheet.getDataRange();
    if ((range.getHeight() < 3) || (range.getWidth() < 3)) {
      // still too small, return empty array
      return new Array();     
    }
  } 
  return range.getValues();
}