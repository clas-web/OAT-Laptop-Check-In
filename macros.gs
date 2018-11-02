//Set the conditional formatting
function conditionFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D3:W46').activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('D3:W45')])
  .whenCellNotEmpty()
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  spreadsheet.getRange('G25').activate();
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(0, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('D3:W45')])
  .whenFormulaSatisfied('=($C3=TRUE)')
  .setBackground('#F4C7C3')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('D3:W45')])
  .whenFormulaSatisfied('=($C3=FALSE)')
  .setBackground('#fbfbfb')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};