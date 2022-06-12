function doGet(e){
  
  var op = e.parameter.action;

  var ss=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1jvdyLkiSJmkNh3owtaHWAcrgNse2DwC54wPoSHAsBRM/edit#gid=0");
  var sheet = ss.getSheetByName("Database");

  
  if(op=="insert")
    return insert_value(e);
   //Logger.log("worked!");
  
}

//Recieve parameter and pass it to function to handle

 


function insert_value(request){
  var ss=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1jvdyLkiSJmkNh3owtaHWAcrgNse2DwC54wPoSHAsBRM/edit#gid=0");
  var sheet = ss.getSheetByName("Database");
 
 
  var f_name = request.parameter.f_name;
  var l_name = request.parameter.l_name;
  var ph_num = request.parameter.ph_num
  var repair_num = request.parameter.repair_num;
  var date = request.parameter.date;
  var make_model = request.parameter.make_model;
  var notes = request.parameter.notes;
  var sku1 = request.parameter.sku1;
  var sku2 = request.parameter.sku2;
  var sku3 = request.parameter.sku3;
  var sku4 = request.parameter.sku4;
  var sku5 = request.parameter.sku5;
  var sku6 = request.parameter.sku6;
  var sku7 = request.parameter.sku7;
  var sku8 = request.parameter.sku8;
  var sku9 = request.parameter.sku9;
  var sku10 = request.parameter.sku10;
  var sku11 = request.parameter.sku11;
  var sku12 = request.parameter.sku12;
  var sku13 = request.parameter.sku13;
  var sku14 = request.parameter.sku14;
  var sku15 = request.parameter.sku15;
  var sku16 = request.parameter.sku16;
  var sku17 = request.parameter.sku17;
  var sku18 = request.parameter.sku18;
  var sku19 = request.parameter.sku19;
  var qty1 = request.parameter.qty1;
  var qty2 = request.parameter.qty2;
  var qty3 = request.parameter.qty3;
  var qty4 = request.parameter.qty4;
  var qty5 = request.parameter.qty5;
  var qty6 = request.parameter.qty6;
  var qty7 = request.parameter.qty7;
  var qty8 = request.parameter.qty8;
  var qty9 = request.parameter.qty9;
  var qty10 = request.parameter.qty10;		
  var qty11 = request.parameter.qty11;		
  var qty12 = request.parameter.qty12;		
  var qty13 = request.parameter.qty13;		
  var qty14 = request.parameter.qty14;		
  var qty15 = request.parameter.qty15;		
  var qty16 = request.parameter.qty16;		
  var qty17 = request.parameter.qty17;		
  var qty18 = request.parameter.qty18;		
  var qty19 = request.parameter.qty19;		


  var flag=1;
  var lr= sheet.getLastRow();
  //add new row with recieved parameter from client
  if(flag==1){
  var d = new Date();
    var currentTime = d.toLocaleString();
  var rowData = sheet.appendRow([currentTime,f_name,l_name,ph_num,repair_num,date,make_model,notes,sku1,sku2,sku3,sku4,sku5,sku6,sku7,sku8,sku9,sku10,sku11,sku12,sku13,sku14,sku15,sku16,sku17,sku18,sku19,qty1,qty2,qty3,qty4,qty5,qty6,qty7,qty8,qty9,qty10,qty11,qty12,qty13,qty14,qty15,qty16,qty17,qty18,qty19]);  
  var result="Insertion successful";
  }
     result = JSON.stringify({
    "result": result
  });  
    
  return ContentService.createTextOutput("Worked").setMimeType(ContentService.MimeType.JAVASCRIPT);  
  

  }
  
