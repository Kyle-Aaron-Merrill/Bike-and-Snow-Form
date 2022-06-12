var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1jvdyLkiSJmkNh3owtaHWAcrgNse2DwC54wPoSHAsBRM/edit#gid=0");

function setWhite(){
  var d_row = SpreadsheetApp.getActive().getActiveRange();
  ss.getRange(d_row.getA1Notation()).setBackground('#FFFFFF');
}

function setBlue() {
  var d_row = SpreadsheetApp.getActive().getActiveRange();
  ss.getRange(d_row.getA1Notation()).setBackground('#4a86e8');
}

function setOrange(){
  var d_row = SpreadsheetApp.getActive().getActiveRange();
  ss.getRange(d_row.getA1Notation()).setBackground('#ff9900');
}

function setGreen(){
  var d_row = SpreadsheetApp.getActive().getActiveRange();
  ss.getRange(d_row.getA1Notation()).setBackground('#00ff00');
}