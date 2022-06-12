function doGet(e){
  
  var op = e.parameter.action;

  var ss=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oSyH3aAoC_qkOxbQ5Uawlhux4bqlIaH90Q7xGdrltwM/edit#gid=0");
  var sheet = ss.getSheetByName("Database");

  
  if(op=="insert")
    return insert_value(e);
   //Logger.log("worked!");
  
}

//Recieve parameter and pass it to function to handle

 


function insert_value(request){
  var ss=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oSyH3aAoC_qkOxbQ5Uawlhux4bqlIaH90Q7xGdrltwM/edit#gid=0");
  var sheet = ss.getSheetByName("Database");
  

  var f_name = request.parameter.f_name;
   var l_name = request.parameter.l_name;
   var ph_num = request.parameter.ph_num;
   var blank = request.parameter.blank;
   var blank2 = request.parameter.blank2;
   var make = request.parameter.make;
   var mdl = request.parameter.mdl;
   var sku_1 = request.parameter.sku_1;
   var sku_2 = request.parameter.sku_2;
   var sku_3 = request.parameter.sku_3;
   var sku_4 = request.parameter.sku_4;
   var sku_5 = request.parameter.sku_5;
   var notes = request.parameter.notes;
   var date_in = request.parameter.date_in;
   var date_out = request.parameter.date_out;
   var ski_snow = request.parameter.ski_snow;
   var stance = request.parameter.stance;
   var weight = request.parameter.weight;
   var height = request.parameter.height;
   var skier_type = request.parameter.skier_type;
   var size = request.parameter.size;
   var boot_make = request.parameter.boot_make;
   var boot_mdl = request.parameter.boot_mdl;
   var bindings = request.parameter.bindings;
   var bootLength = request.parameter.bootLength;
   var age = request.parameter.age;
  		


  var flag=1;
  var lr= sheet.getLastRow();
  //add new row with recieved parameter from client
  if(flag==1){
  var d = new Date();
    var currentTime = d.toLocaleString();
  var rowData = sheet.appendRow([currentTime,f_name,l_name,ph_num,blank,blank2,make,mdl,sku_1,sku_2,sku_3,sku_4,sku_5,notes,date_in,date_out,ski_snow,stance,weight,height,skier_type,size,boot_make,boot_mdl,bindings,bootLength,age]);  
  var result="Insertion successful";
  }
     result = JSON.stringify({
    "result": result
  });  
    
  return ContentService.createTextOutput("Worked").setMimeType(ContentService.MimeType.JAVASCRIPT);  
  

  }
  
