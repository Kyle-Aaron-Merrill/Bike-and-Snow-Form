function refresh(){
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1jvdyLkiSJmkNh3owtaHWAcrgNse2DwC54wPoSHAsBRM/edit#gid=0");
  var data = ss.getSheetByName("Database");
  var rep = ss.getSheetByName("Repairs");
  var config = ss.getSheetByName("Config");
  var labor = ss.getSheetByName("Labor");
  var parts = ss.getSheetByName("Parts");
  var ui = SpreadsheetApp.getUi();
  var priceArr = [];
  var done = false;

  //config.getRange(2,2).setValue(1);
  

  //rep.clear();
  updateRow();
  //delRow();
  var d_row = config.getRange(1,2).getValue();
  start();

function start(){
  

  while (done == false){
    buildDb();
    buildRep();
  } 

}
function buildRep(){
  createTemplate();
  addData();
  
}

function buildDb(){
  clearRows();
  moveSku();
  linkCell();
}
function delRow(){
  var del = "delete";
  d_row = whatRow(del);
  data.deleteRow(d_row);
  var row = config.getRange(1,2).getValue();
  config.getRange(1,2).setValue(row--);
  start();
}
function updateRow(){
  var d_row = SpreadsheetApp.getActive().getActiveRange().getRow();
  var row = ((d_row -3) * 28) + 1;
  var end = config.getRange(2,2).getValue();
  for(var i = row; i<end;i++){
    rep.getRange(i,1,27,6).clear();
  }
  config.getRange(2,2).setValue(row);
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


function createTemplate(){
  var slip_row = config.getRange(2,2).getValue();

  var conditionalFormatRules = rep.getConditionalFormatRules();
  rep.setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = rep.getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule().setRanges([rep.getRange(slip_row,1,6,6)]) .whenCellEmpty().setBackground('#EA9999').build());
        
  rep.getRange(slip_row,1).setValue("NAME").setHorizontalAlignment('right').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,3).setValue("DATE").setHorizontalAlignment('right').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,5).setValue("REPAIR#").setHorizontalAlignment('right').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 1,1).setValue("PHONE").setHorizontalAlignment('right').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 2,5).setValue("QTY").setHorizontalAlignment('center').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 2,4).setValue("SKU").setHorizontalAlignment('center').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 2,1,1,3).setValue("DESCRIPTION").setHorizontalAlignment('center').setBackground('#fff2cc').mergeAcross().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 2,6).mergeAcross().setValue("CO$T").setHorizontalAlignment('center').setBackground('#fff2cc').mergeAcross().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 22,1,1,5).setValue("TOTAL").mergeAcross().setHorizontalAlignment('right').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 23,1).setValue("BIKE").setHorizontalAlignment('right').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 24,1).setValue("NOTES").setHorizontalAlignment('right').setVerticalAlignment('top').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);    
  Logger.log("Template created"); 
}

function moveSku(){
  for (var i = 9; i < 28; i++){
    var sku_check = data.getRange(d_row,i).getValue();
    if(sku_check == ""){
      for (var x = i + 1; x < 28; x++){
        sku_check = data.getRange(d_row,x).getValue();
        if(sku_check != ""){
          moveQty(x,i);
          data.getRange(d_row,i).setValue(sku_check);
          data.getRange(d_row,x).setValue("");
          break;
        }
      } 
    }
  }
}
function moveQty(x,i){
  var n = x + 19;
  var qty_check = data.getRange(d_row,n).getValue();
  if(qty_check != ""){
    data.getRange(d_row,i+19).setValue(qty_check);
    data.getRange(d_row,n).setValue("");
  }
}

function checkRow(){
  var checker = data.getRange(d_row + 1,1).getValue();
  if (checker == null || checker == '' || checker == ""){
    Logger.log(d_row);
    done = true;
    config.getRange(1,2).setValue(d_row);
  }
}

function clearRows(){
  var del_row = d_row;
  var done = false;

  while(done == false){
        //declaring variables that will change inside while loop
        // @ts-ignore
    var date_checker = data.getRange(del_row+1,1).getValue();
        // @ts-ignore
    var name_checker = data.getRange(del_row+1,2).getValue();
        
    if (date_checker != "" && name_checker == ""){ 
        //deletes oopsie inputs
        // @ts-ignore
      data.deleteRow(del_row+1);
      }
      else done = true;
      }
}

function addData(){

  var fName = data.getSheetValues(d_row,2,1,1);
  var lName = data.getSheetValues(d_row,3,1,1);
  var repair_num = data.getSheetValues(d_row,5,1,1);
  var phNum = data.getSheetValues(d_row,4,1,1);
  var date = data.getSheetValues(d_row,6,1,1);
  var make = data.getSheetValues(d_row,7,1,1);
  var notes = data.getSheetValues(d_row,8,1,1);

  var slip_row = config.getRange(2,2).getValue();
  var name = fName + " " + lName;
  rep.getRange(slip_row,2).setValue(name).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,4).setValue(date).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,6).setValue(repair_num).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.autoResizeColumn(6);
  rep.getRange(slip_row + 1, 2,1,5).mergeAcross().setHorizontalAlignment('center').setValue(phNum).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 23,2,1,5).mergeAcross().setValue(make).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 24,2,1,5).mergeAcross().setValue(notes).setWrap(true).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
 addSkus(slip_row);
}

function addSkus(row){
  row += 3;
  for (col = 9; col < 28; col++){
    Logger.log(data.getRange(d_row,col).getValues());
    var sku = data.getRange(d_row,col).getValues();
    var qty = qtyChecker(col);
    var desc = null;
    var price = null;
    if(sku != "" || sku > 0){
      if(sku != "" || sku > 0){
        var obj = null;
        obj = laborCheck(sku);
        if (obj == null){
          obj = partCheck(sku);
          if(obj == null){
            obj = addSkuTool(sku);
          }
        }
      }
      if(obj != null){
        sku = obj[0];
        desc = obj[1];
        price = obj[2];
        price = calc(price,qty);

        insert(sku,desc,price,qty,row);
        row = nextLine(row);
      }
    }
    else{
      row = nextLine(row);
    }
  }
  var total = totalCalc();
  rep.getRange(row,6).setValue(total).setBackground('#C41616');
  nextRow(row);
}

function linkCell(){
  var row = config.getRange(2,2).getValue();
  var cell = data.getRange(d_row,5);
  var range = rep.getRange(row,1,25,6);
  var repairNum = config.getRange(4,2).getValue();
  var invoiceNum = data.getRange(d_row,48).getValue().toLowerCase();

  if(invoiceNum != "" && invoiceNum != null){
    repairNum = invoiceNum;
  }
  else{
    repairNum = repairNum + 1;
    config.getRange(4,2).setValue(repairNum);
  }

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(repairNum)
    .setLinkUrl('#gid=' + rep.getSheetId() + '&range=' + range.getA1Notation())
    .build();
    cell.setRichTextValue(richText);


}
function totalCalc(){
  var total = 0;
  for(var i = 0; i < priceArr.length;i++){
    if(priceArr[i] != null){
    total += Number(priceArr[i]);
    }
  }
  priceArr.splice(0);
  return total;
  
}

function nextRow(row){
  checkRow();
  d_row++;
  config.getRange(2,2).setValue(row + 6);
}

function calc(price,qty){
  return (price * qty).toFixed(2)
}

function qtyChecker(column){
  var qty = data.getRange(d_row,column + 19).getValue();
  if (qty < 1){qty = 1}
  return qty;
}

function addSkuTool(sku){
  var part_cnt = parts.getMaxRows();
  var obj = null;
  var desc = descTool(sku);
  var price = priceTool(sku);

  
  parts.insertRowAfter(part_cnt);
  part_cnt = parts.getMaxRows();

  parts.getRange(part_cnt,1).setValue(sku).setBackground('#a2c4c9').setFontColor('black');
  parts.getRange(part_cnt,3).setValue(price).setBackground('#a2c4c9').setFontColor('black');
  parts.getRange(part_cnt,2).setValue(desc).setBackground('#980000');

  price = price *1.06;
  obj = [sku,desc,price];

  return obj;
}

function descTool(sku){
  var response = ui.prompt('SKU: ' + sku + ' cannot be found', 'Add description below', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var desc = response.getResponseText();
  }
  else{
    var desc = 'Add Description';
  }
  return desc;
}

function priceTool(sku){
  var response = ui.prompt('SKU: ' + sku + ' cannot be found', 'Add Price below', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var price = Number(response.getResponseText());
  }
  else{
    var price = 'Add Price';
  }
  return price;
}

function laborCheck(sku){
  var labor_cnt = labor.getMaxRows();
  var labor_data = labor.getRange(1,1,labor_cnt,3).getValues();
  var obj = null;
  for (var i = 1; i < labor_cnt; i++){
    var check = labor_data[i][0];
    if(check == sku){
       if(sku == '5574997'){
        var name = data.getRange(d_row,2).getValue() + " " + data.getRange(d_row,3).getValue();
        var miscPrice = data.getRange(d_row,47).getValue();
        desc = 'Misc Labor';
        price = miscPriceChecker(miscPrice,name);
      }
      else{
      var desc = labor_data[i][1];
      var price = labor_data[i][2];
      }
      obj = [sku,desc,price];
      break;
    }
  }
  return obj;
}

function partCheck(sku){
  var part_cnt = parts.getMaxRows();
   var part_data = parts.getRange(1,1,part_cnt,3).getValues();
  var obj = null;
  for (var i = 1; i < part_cnt; i++){
    var check = part_data[i][0];
    if(check == sku){
      var desc = part_data[i][1];
      var price = part_data[i][2] * 1.06;
      obj = [sku,desc,price];
      break;
    }
  }
  return obj;
}

function nextLine(row){
  rep.getRange(row,6).setBackground('#e06666');
  rep.getRange(row,5).setBackground('#f6b26b');
  rep.getRange(row,4).setBackground('#93c47d');
  rep.getRange(row,1,1,3).mergeAcross().setBackground('#76a5af');
  row++;
  return row;
}

function insert(sku,desc,price,qty,row){
  rep.getRange(row,6).setValue(price).setBackground('#e06666');
  rep.getRange(row,5).setValue(qty).setBackground('#f6b26b');
  rep.getRange(row,4).setValue(sku).setBackground('#93c47d');
  rep.getRange(row,1,1,3).mergeAcross().setValue(desc).setBackground('#76a5af');
  priceArr.push(price);
}

function miscPriceChecker(miscPrice,name){
  if(miscPrice < 1){
    var response = ui.prompt('Enter a misc price for ' + name + ' below', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var price = Number(response.getResponseText());
      data.getRange(d_row,47).setValue(price);
    }
    else{
      Logger.log('User canceled');
      var price = 20;
      data.getRange(d_row,47).setValue(price);
    }
  }
  if (miscPrice > 1){
      var price = miscPrice;
      data.getRange(d_row,47).setValue(price);
  }  
    return price;
    
}

}