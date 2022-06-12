function refresh(){
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oSyH3aAoC_qkOxbQ5Uawlhux4bqlIaH90Q7xGdrltwM/edit#gid=0");
  var data = ss.getSheetByName("Database");
  var rep = ss.getSheetByName("Service");
  var config = ss.getSheetByName("Config");
  var labor = ss.getSheetByName("Labor");
  var parts = ss.getSheetByName("Parts");
  //var ui = SpreadsheetApp.getUi();
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
  setDin();
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
  var response = SpreadsheetApp.getUi().prompt('Please enter a row to ' + type + ': ');

    if (response.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
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
        
  rep.getRange(slip_row,1).setValue("NAME").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,3).setValue("DATE IN").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,5).setValue("DATE OUT").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 1,1).setValue("PHONE").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 1,4).setValue("REPAIR #").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 2,1).setValue("DIN").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 3,1).setValue("EQUIPMENT").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 3,4).setValue("STANCE").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 4,1).setValue("BOARD").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 4,3).setValue("MODEL").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 4,5).setValue("SIZE (cm)").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 5,1).setValue("BOOTS").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 5,3).setValue("MODEL").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 5,5).setValue("BINDINGS").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 6,1).setValue("WEIGHT").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 6,3).setValue("HEIGHT").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 6,5).setValue("TYPE").setBackground('#D27FFC').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 7,5).setValue("QTY").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 7,4).setValue("SKU").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 7,1,1,3).merge().setValue("DESCRIPTION").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 7,6).setValue("CO$T").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 13,1,1,5).merge().setValue("TOTAL").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 14,1).setValue("MAKE/MODEL").setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 15,1).setValue("NOTES").setVerticalAlignment('top').setBackground('#fff2cc').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);    
  Logger.log("Template created"); 
}

function moveSku(){
  for (var i = 9; i < 14; i++){
    var sku_check = data.getRange(d_row,i).getValue();
    if(sku_check == ""){
      for (var x = i + 1; x < 14; x++){
        sku_check = data.getRange(d_row,x).getValue();
        if(sku_check != ""){
          //moveQty(x,i);
          data.getRange(d_row,i).setValue(sku_check);
          data.getRange(d_row,x).setValue("");
          break;
        }
      } 
    }
  }
}
function moveQty(x,i){
  var n = x + 14;
  var qty_check = data.getRange(d_row,n).getValue();
  if(qty_check != ""){
    data.getRange(d_row,i+19).setValue(qty_check);
    data.getRange(d_row,n).setValue("");
  }
}

function setDin(){
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oSyH3aAoC_qkOxbQ5Uawlhux4bqlIaH90Q7xGdrltwM/edit#gid=0");
  var data = ss.getSheetByName("Database");
  var rep = ss.getSheetByName("Service");
  var config = ss.getSheetByName("Config");
  var labor = ss.getSheetByName("Labor");
  var parts = ss.getSheetByName("Parts");
  //var ui = SpreadsheetApp.getUi();
  var priceArr = [];
  var done = false;

  

  var weight = data.getRange(d_row,19).getValue();
  var height = data.getRange(d_row,20).getValue();
  var type = data.getRange(d_row,21).getValue();
  var bootLength = data.getRange(d_row,26).getValue();
  var age = data.getRange(d_row,27,1,1).getValue();

    //Array for din calculator
    dinArr = [[0.75,0.75,0.75," "," "," "," "," "],
    [1.00,0.75,0.75,0.75," "," "," "," "],
    [1.50,1.25,1.25,1.00," "," "," "," "],
    [2.00,1.75,1.50,1.50,1.25," "," "," "],
    [2.50,2.25,2.00,1.75,1.50,1.50," "," "],
    [3.00,2.75,2.50,2.25,2.00,1.75,1.75," "],
    [" ",3.5,3.0,2.75,2.5,2.25,2.00," "], 
    [" "," ",3.50,3.0,3.0,2.75,2.5," "],
    [" "," ",4.5,4.0,3.5,3.5,3.0," "],
    [" "," ", 5.5,5,4.5,4.0,3.5,3.0],
    [" "," ",6.5,6.0,5.5,5.0,4.5,4.0],
    [" "," ",7.5,7.0,6.5,6.0,5.5,5.0],
    [" "," "," ",8.5,8.0,7.0,6.5,6.0],
    [" "," "," ",10.0,9.5,8.5,8.0,7.5],
    [" "," "," ",11.5,11.0,10.0,9.5,9.0],
    [" "," "," "," "," ",12.0,11.0,10.5]];

    //Array for twist torque
    twistArr = [5,8,11,14,17,20,23,27,31,37,43,50,58,67,78,91,105,121,137];
    //Array for forward lean torque
    forwardArr = [18,29,40,52,64,75,87,102,120,141,165,194,229,271,320,380,452,520,588];

    //checks if all parameters are met
    if(weight != "" && height != "" && age != "" && bootLength != ""){
      var weightDin = weightCalc();
      var heightDin = heightCalc();

      if(weightDin <= heightDin) {
        var dinRow = weightDin;
      }
      else {
        var dinRow = heightDin;
      }

      var n = typeCalc();
      dinRow = dinRow + n;

      var dinCol = bootCalc();
      var n = ageCalc();

      dinRow = dinRow - n;
      setDin();

    }
    else{
      data.getRange(d_row,6).setValue("Not enough info to calculate din")
    }
    //functions for dinCalc
    function weightCalc(){
      //setting din to 0
      var din = 0;

      if(weight > 21  && weight < 30){
        din = 1;
      }
      if(weight > 29  && weight < 39){
        din = 2;
      }
      if(weight > 38  && weight < 48){
        din = 3;
      }
      if(weight > 47  && weight < 57){
        din = 4;
      }
      if(weight > 56  && weight < 67){
        din = 5;
      }
      if(weight > 66  && weight < 79){
        din = 6;
      }
      if(weight > 78  && weight < 92){
        din = 7;
      }
      if(weight > 91  && weight < 108){
        din = 8;
      }
      if(weight > 107  && weight < 126){
        din = 9;
      }
      if(weight > 125  && weight < 148){
        din = 10;
      }
      if(weight > 147  && weight < 175){
        din = 11;
      }
      if(weight > 174  && weight < 210){
        din = 12;
      }
      if(weight > 209){
        din = 13;
      }
      //returns current row of array
      return din;
    }
    function heightCalc(){
    //sets din to 7
    var din = 8;

    if(height == "4,10"){
      din = 8;
    }
    if(height == "4,11-5,1"){
      din = 9;
    }
    if(height == "5,2-5,5"){
      din = 10;
    }
    if(height == "5,6-5,10"){
      din = 11;
    }
    if(height == "5,11-6,4"){
      din = 12;
    }
    if(height == "6,5"){
      din = 13;
    }
    //returns current row of array
    return din
    }
    function typeCalc(){
      var t = 0;

      if(type == 1){
        t = 0;
      }
      if(type == 2){
        t = 1;
      }
      if(type == 3 && age > 9){
        t = 2;
      }
      else if (type == 3 && age <= 9){
        t = 2;
      }
      return t;
    }
    function bootCalc(){
      var din = 1;
      if(bootLength <= 230){
        din = 1;
      }
      if(bootLength >= 231 && bootLength <= 250){
        din = 2;
      }
      if(bootLength >= 251 && bootLength <= 270){
        din = 3;
      }
      if(bootLength >= 271 && bootLength <= 290){
        din = 4;
      }
      if(bootLength >= 291 && bootLength <= 310){
        din = 5;
      }
      if(bootLength >= 311 && bootLength <= 330){
        din = 6;
      }
      if(bootLength >= 331 && bootLength <= 350){
        din = 7;
      }
      if(bootLength >= 351){
        din = 8;
      }
      return din;
    }
    function ageCalc(){
      var t = 0;
      if(age <= 9 || age >= 50){
      t = 1;
      }
      else t = 0;

      return t;
    }
    function setDin(){
    //change this number if torque rows change  
    var i=2;
    //if (dinRow == 6){dinCol = dinCol+1;}
    //if (dinRow > 6 && dinRow < 12){dinCol = dinCol+1;} 
    //if (dinRow > 11 && dinRow < 15) {dinCol = dinCol+2;} 
    //if (dinRow == 15) {dinCol = dinCol+4;}
    var twistRow = dinRow;
    var forwardRow = dinRow;
    Logger.log(dinArr[dinRow-1][dinCol-1])
    console.log(" Din: " + dinArr[dinRow-1][dinCol-1]+ "\n" + " Twist Range: " + twistArr[twistRow-i] + "nm - " + twistArr[twistRow+ i] + "nm \n" + " Forward Lean Range: " + forwardArr[forwardRow- i] + "nm - " + forwardArr[forwardRow+ i] + "nm");
    data.getRange(d_row,6).setValue(" Din: " + dinArr[dinRow-1][dinCol-1]+ "\n" + " Twist Range: " + twistArr[twistRow- i] + "nm - " + twistArr[twistRow+ i] + "nm \n" + " Forward Lean Range: " + forwardArr[forwardRow- i] + "nm - " + forwardArr[forwardRow+ i] + "nm");
    
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
  var dateIn = data.getSheetValues(d_row,15,1,1);
  var dateOut = data.getSheetValues(d_row,16,1,1);
  var din = data.getSheetValues(d_row,6,1,1)
  var make = data.getSheetValues(d_row,7,1,1);
  var mdl = data.getSheetValues(d_row,8,1,1);
  var ski_snow = data.getSheetValues(d_row,17,1,1);
  var stance = data.getSheetValues(d_row,18,1,1);
  var bootMake = data.getSheetValues(d_row,23,1,1);
  var bootModel = data.getSheetValues(d_row,24,1,1);
  var bidings = data.getSheetValues(d_row,25,1,1);
  var miscLabor = data.getSheetValues(d_row,28,1,1);
  var notes = data.getSheetValues(d_row,14,1,1);
  var weight = data.getSheetValues(d_row,19,1,1);
  var height = data.getSheetValues(d_row,20,1,1);
  var type = data.getSheetValues(d_row,21,1,1);
  var size = data.getSheetValues(d_row,22,1,1);
  var bootLength = data.getSheetValues(d_row,26,1,1);
  var age = data.getSheetValues(d_row,27,1,1);

  var slip_row = config.getRange(2,2).getValue();
  var name = fName + " " + lName;
  rep.getRange(slip_row,2).setValue(name).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,4).setValue(dateIn).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row,6).setValue(dateOut).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 1,2,1,2).merge().setValue(phNum).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 1,5,1,2).merge().setValue(repair_num).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 2,2,1,5).merge().setHorizontalAlignment('center').setValue(din).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 3,5,1,2).merge().setValue(stance).setHorizontalAlignment('left').setBackground('#92DC80').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 3,2,1,2).merge().setValue(ski_snow).setHorizontalAlignment('left').setBackground('#92DC80').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);


  rep.getRange(slip_row + 4,2,1,1).setValue(make).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 4,4,1,1).setValue(mdl).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 4,6,1,1).setValue(size).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 5,2,1,1).setValue(bootMake).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 5,4,1,1).setValue(bootModel).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 5,6,1,1).setValue(bidings).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 6,2,1,1).setValue(weight).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 6,4,1,1).setValue(height).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 6,6,1,1).setValue(type).setBackground('#FFB141').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 14,2,1,5).merge().setValue(make + " " + mdl).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rep.getRange(slip_row + 15,2,1,5).merge().setValue(notes).setBackground('#94D2E6').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
 addSkus(slip_row);
}

function addSkus(row){
  row += 8;
  for (col = 9; col <= 13; col++){
    //Logger.log(data.getRange(d_row,col,1,1) + " Row: " + d_row + " Column: " + col);
    var sku = data.getRange(d_row,col).getValue();
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
  var range = rep.getRange(row,1,16,6);
  var repairNum = config.getRange(4,2).getValue();
  
  repairNum = repairNum + 1;
  config.getRange(4,2).setValue(repairNum);

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(repairNum)
    .setLinkUrl('#gid=' + rep.getSheetId() + '&range=' + range.getA1Notation())
    .build();
    cell.setRichTextValue(richText );


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
  if (sku >= 0){
    var desc = descTool(sku);
    var price = priceTool(sku);

    
    parts.insertRowAfter(part_cnt);
    part_cnt = parts.getMaxRows();

    parts.getRange(part_cnt,1).setValue(sku).setBackground('#a2c4c9').setFontColor('black');
    parts.getRange(part_cnt,3).setValue(price).setBackground('#a2c4c9').setFontColor('black');
    parts.getRange(part_cnt,2).setValue(desc).setBackground('#980000');

    price = price *1.06;
    obj = [sku,desc,price];
  }
  else{
    var desc = sku;
    sku = "NULL";
    var price = "NULL";
    obj = [sku,desc,price];
  }

  return obj;
}

function descTool(sku){
  var response = SpreadsheetApp.getUi().prompt('SKU: ' + sku + ' cannot be found', 'Add description below', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
    var desc = response.getResponseText();
  }
  else{
    var desc = 'Add Description';
  }
  return desc;
}

function priceTool(sku){
  var response = SpreadsheetApp.getUi().prompt('SKU: ' + sku + ' cannot be found', 'Add Price below', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
    var price = Number(response.getResponseText());
  }
  else{
    var price = 'Add Price';
  }
  return price;
}

function email(){
  var response = ui.prompt("Enter the row you want to ship to BHBS",ui.ButtonSet.OK_CANCEL);
  var d_row = Number(response.getResponseText());

  var workObj = new Array;
  for (var i=9;i<=13;i++){
    var sku = data.getRange(d_row,i).getValue()
     if(sku == '4234097' || sku == '7524872' || sku == '4251981' || sku == '7524869'){
        var name = data.getRange(d_row,2).getValue() + " " + data.getRange(d_row,3).getValue();
        var obj = laborCheck(sku);
        workObj.push(String(obj[0]+" "+obj[1]));
        Logger.log(workObj)
      }
  }
  var string = stringBuilder(workObj);
  bhbsShipTool(name, string);
}

function bhbsShipTool(custy, text){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Customer: " + custy + " requested work in Birmingham. When able too, add a tracking number below.", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL)
  var tracking = response.getResponseText();
  Logger.log(tracking.length)
    if (tracking != "" && tracking != " " && tracking.length <= 15 && tracking.length >=12 && response.getSelectedButton() == ui.Button.OK){ 
      var equipment = data.getRange(d_row,7).getValue() + " " + data.getRange(d_row,8).getValue();
      var subj = "Local Skis in your area wanna meet";
      var str = "Hey Freds,\nkeep an eye out for " + custy + "'s " + equipment + "\n\nThey requested:\n(" + text + "\n Here is the tracking number: " + tracking + ".\n - Automated message from King Fred";
      //Logger.log(str);
      var email = config.getRange(5,2).getValue();
      MailApp.sendEmail(email,subj,str);
    }
    else if (tracking || "" && tracking || " " && tracking.length >= 15 || tracking.length <=12 && response.getSelectedButton() == ui.Button.OK){
      ui.alert("Tracking number isnt correct");
    }
    else if (response.getSelectedButton() == ui.Button.CANCEL){
      Logger.log("closed UI");
    }
}

function stringBuilder(workObj){
  var string = "";
  for (var k=0;k<workObj.length;k++){
    string += String("("+workObj[k] + ")\n");
  } 
  Logger.log(string);
  return string;
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
      if(sku == '4234097' || sku == '7524872' || sku == '4251981' || sku == '7524869'){
        data.getRange(d_row,1,1,28).setBackground("#FC00CE")
        var desc = labor_data[i][1];
        var price = labor_data[i][2];
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
  rep.getRange(row,1,1,3).merge().setBackground('#76a5af');
  row++;
  return row;
}

function insert(sku,desc,price,qty,row){
  rep.getRange(row,6).setValue(price).setBackground('#e06666');
  rep.getRange(row,5).setValue(qty).setBackground('#f6b26b');
  rep.getRange(row,4).setValue(sku).setBackground('#93c47d');
  rep.getRange(row,1,1,3).merge().setValue(desc).setBackground('#76a5af');
  priceArr.push(price);
}

function miscPriceChecker(miscPrice,name){
  if(miscPrice < 1){
    var response = SpreadsheetApp.getUi().prompt('Enter a misc price for ' + name + ' below', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
      var price = Number(response.getResponseText());
      data.getRange(d_row,28).setValue(price);
    }
    else{
      Logger.log('User canceled');
      var price = 20;
      data.getRange(d_row,28).setValue(price);
    }
  }
  if (miscPrice > 1){
      var price = miscPrice;
      data.getRange(d_row,28).setValue(price);
  }  
    return price;
    
}
}