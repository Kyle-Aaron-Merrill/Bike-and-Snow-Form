  function clone() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oSyH3aAoC_qkOxbQ5Uawlhux4bqlIaH90Q7xGdrltwM/edit#gid=0");
  var rep = ss.getSheetByName("Service");
  var date = new Date();
  var mt = date.getMonth();
  var yr = date.getFullYear();
  var months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  var config = ss.getSheetByName("Config");
  var data = ss.getSheetByName("Database");
  var done = "false";
  //var j = config.getRange(4,2).getValue();
  
  getlastMonth();
  if (mt == 11){
    yr = 2021
  }
  var currentD =  months[mt]+ " "  + yr; 
  Logger.log(currentD);
  
  checkSheet();

  function exp() { 
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1rXNws0pPzn_Zix3W53LCtb164L5pyTSnyDz7JZWoPcc/edit#gid=0"); 
  var data = ss.getSheetByName("Database");
  var folders = DriveApp.getFolders();

    while (folders.hasNext()) {
      var folder = folders.next();
      if(folder == "Ski_DB"){
        Logger.log(folder.getName() + " found");
        var parentFolder = folder.getId();
        Logger.log(parentFolder.toString());
        yearFolder(parentFolder);
        break;
      }
    }
  }
  function refreshAll(){
    var start = 3;
    var end = getEnd(start);
    config.getRange(1,2).setValue(end);
  }
  function getEnd(row){
    var done = false;
    while(done == false){
      var check = data.getRange(row,1).getValue();
      if(check == ""){
        var end = row;
        done = true;
      }
      row++;
    }
    return end;
  }
  function checkSheet(){
    var currentSheet = ss.getSheetByName(currentD);
    if(currentSheet == null){
      refreshAll();
      exp();
      dataBackup();
      deleteData();
    }
  }
  function yearFolder(parentFolder){
    var subFolder = DriveApp.getFolderById(parentFolder).getFolders();

    if (subFolder.hasNext() == false){
      DriveApp.getFolderById(parentFolder).createFolder(yr);
    }
    
    subFolder = DriveApp.getFolderById(parentFolder).getFolders();

    while (subFolder.hasNext() == true) {
    var folder = subFolder.next();
      Logger.log(folder.toString());
      if(folder == yr){
        Logger.log(folder.getName() + " found");
        var yearFolder = folder.getId();
        //Logger.log(yearFolder.toString());
        break;
      }
      else if (folder != yr){
        Logger.log(folder.getName() + " not found found. Creating a new folder");
        DriveApp.getFolderById(parentFolder).createFolder(yr);
        var yearFolder = folder.getId();
        Logger.log(yearFolder.toString());
        break;
      }
    }
    addFile(yearFolder);
  }
  function addFile(yearFolder){
    var fileLoc = DriveApp.getFolderById(yearFolder);

    fileLoc.createFile(currentD,fileSave(),MimeType.CSV);
  }

  function fileSave(){
    var csv = "Timestamp,First Name,Last Name,Phone Number,Repair Number,Din,Make,Model,sku 1,sku 2,sku 3,sku 4,sku 5,Notes,Date in,Date out,Ski,Snow,weight,height,Skier Type,Size,Boot make,Boot model,Bindings,Boot Length,Age,Misc Price"+ "\r\n";

    var start = config.getRange(3,2).getValue();
    var end = config.getRange(1,2).getValue();
    var maxCol = data.getMaxColumns();

    for(var row = start; row <= end; row++){
      for(var col = 1; col < maxCol; col++){
        var dat = data.getRange(row,col).getValue().toString();
        //Logger.log(dat);
        if(dat.indexOf(",") != -1 || dat.indexOf("\n") != -1){
          csv += "\"" + dat + "\"" + ",";
          Logger.log("\"" + dat + "\"" + ",");
        }
        //change depending on time zone
    
        else{
        csv += dat + ",";
        }
      }
      csv += "\r\n";
    }
    return csv;
  }
  function dataBackup(){
    var start = config.getRange(3,2).getValue();
    var end = config.getRange(1,2).getValue();

    var lastMonth = months[getlastMonth()].toString() + " " + yr.toString();
    var oldSheet = ss.getSheetByName(lastMonth);

    if(oldSheet != null){
    ss.deleteSheet(oldSheet);
    }

    ss.insertSheet(currentD);
    var sheet = ss.getSheetByName(currentD);
    ss.moveActiveSheet(6);

    var topOfSheet = data.getRange(2,1,1,28).getValues();
    sheet.getRange(1,1,1,28).setValues(topOfSheet); 
    
    var i = 2;

    for(var row = start; row <= end; row++){
      var value = data.getRange(row,1,1,28).getValues();
      var color = data.getRange(row,2).getBackground();
      sheet.getRange(i,1,1,28).setValues(value).setBackground(color);
      i++;
    }
  }
  function getlastMonth(){
    mt--;
    if(mt == -1){
      mt = 11;
    }
    return mt;
  }
  function deleteData(){
     var start = config.getRange(3,2).getValue();
     var end = config.getRange(1,2).getValue();

     data.deleteRows(start--,end-start++);

    config.getRange(3,2).setValue(start);
    config.getRange(1,2).setValue(start);  
  }
}



