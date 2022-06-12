var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oSyH3aAoC_qkOxbQ5Uawlhux4bqlIaH90Q7xGdrltwM/edit#gid=0");

function setWhite() {
  if (SpreadsheetApp.getActiveSheet() != null){
  var d_row = SpreadsheetApp.getActive().getActiveRange().getA1Notation();
  ss.getRange(d_row).setBackground('#FFFFFF');
  }
}
function setBlue(){
  if (SpreadsheetApp.getActiveSheet() != null){
  var d_row = SpreadsheetApp.getActive().getActiveRange().getA1Notation();
  ss.getRange(d_row).setBackground('#4a86e8');
  }
}
function setOrange(){
  if (SpreadsheetApp.getActiveSheet() != null){
  var d_row = SpreadsheetApp.getActive().getActiveRange().getA1Notation();
  ss.getRange(d_row).setBackground('#ff9900');
  }
}
function setGreen(){
  if (SpreadsheetApp.getActiveSheet() != null){
  var d_row = SpreadsheetApp.getActive().getActiveRange().getA1Notation();
  ss.getRange(d_row).setBackground('#00ff00');
  }
}
function setPink(){
  if (SpreadsheetApp.getActiveSheet() != null){
  var d_row = SpreadsheetApp.getActive().getActiveRange().getA1Notation();
  ss.getRange(d_row).setBackground('#FC00CE');
  }
}

