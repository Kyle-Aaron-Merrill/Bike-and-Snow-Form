function deleteRow(){
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1jvdyLkiSJmkNh3owtaHWAcrgNse2DwC54wPoSHAsBRM/edit#gid=0");
  var config = ss.getSheetByName("Config");
  var data = ss.getSheetByName("Database");
  var rep = ss.getSheetByName("Repairs");
  var ui = SpreadsheetApp.getUi();

  delRow();

  function delRow(){
    var del = "delete";
    var d_row = whatRow(del);
    data.deleteRow(d_row);

    var row = ((d_row -3) * 28) + 1;
    var end = row + 28;
    for(var i = row; i<end;i++){
      rep.deleteRow(row);
      rep.insertRowAfter(rep.getMaxRows());
    }
    config.getRange(2,2).setValue(config.getRange(2,2).getValue() - 28);
    config.getRange(1,2).setValue(d_row);
  }
  function whatRow(type){
    var response = ui.prompt('Please enter a row to ' + type + ': ');

      if (response.getSelectedButton() == ui.Button.OK) {
        var d_row = (response.getResponseText());
      }   
      else {
        var d_row = 0;
      }
      return Number(d_row);
  }
}