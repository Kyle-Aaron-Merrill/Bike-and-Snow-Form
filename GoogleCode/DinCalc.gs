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

  var d_row = 138;

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