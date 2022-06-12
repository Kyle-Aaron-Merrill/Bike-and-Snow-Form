
  
  // Make an AJAX call to Google Script
  function insert_bike() {
    var script_url = "https://script.google.com/macros/s/AKfycbzKJT_aznOXPrOC8xL1_GQdbUAHTd6KIRINjZTscexqHtAPKqCpKzcNRv1l6g4Ho8lreQ/exec";

    
    var f_name=	$("#bike_f_name").val();
	var l_name= $("#bike_l_name").val();
    var ph_num= $("#bike_ph_num").val();
    var repair_num= $("#blank").val();
    var make= $("#bike_make").val();
    var model= $("#bike_model").val();
    var size= $("#bike_size").val();
    var notes= $("#bike_notes").val();
    var sku1= $("#bike_sku_1").val();            
    var sku2= $("#bike_sku_2").val();     
    var sku3= $("#bike_sku_3").val();  
    var sku4= $("#bike_sku_4").val();  
    var sku5= $("#bike_sku_5").val();  
    var sku6= $("#bike_sku_6").val();  
    var sku7= $("#bike_sku_7").val();  
    var sku8= $("#bike_sku_8").val();  
    var sku9= $("#bike_sku_9").val();  
    var sku10= $("#bike_sku_10").val();
    var sku11= $("#bike_sku11").val();
    var sku12= $("#bike_sku12").val();
    var sku13= $("#bike_sku13").val();
    var sku14= $("#bike_sku14").val();
    var sku15= $("#bike_sku15").val();
    var sku16= $("#bike_sku16").val();
    var sku17= $("#bike_sku17").val();
    var sku18= $("#bike_sku18").val();
    var sku19= $("#bike_sku19").val();
    var sku20= $("#bike_sku20").val();
	  var qty1= $("#bike_qty_1").val();
	  var qty2= $("#bike_qty_2").val();
    var qty3= $("#bike_qty_3").val();
    var qty4= $("#bike_qty_4").val();
    var qty5= $("#bike_qty_5").val();
    var qty6= $("#bike_qty_6").val();
    var qty7= $("#bike_qty_7").val();
    var qty8= $("#bike_qty_8").val();
    var qty9= $("#bike_qty_9").val();
    var qty10= $("#bike_qty_10").val();
    var qty11= $("#bike_qty11").val();
    var qty12= $("#bike_qty12").val();
    var qty13= $("#bike_qty13").val();
    var qty14= $("#bike_qty14").val();
    var qty15= $("#bike_qty15").val();
    var qty16= $("#bike_qty16").val();
    var qty17= $("#bike_qty17").val();
    var qty18= $("#bike_qty18").val();
    var qty19= $("#bike_qty19").val();
    var qty20= $("#bike_qty20").val();


    var url = script_url+"?callback=ctrlq&qty20="+qty20+"&qty19="+qty19+"&qty18="+qty18+"&qty17="+qty17+"&qty16="+qty16+"&qty15="+qty15+"&qty14="+qty14+"&qty13="+qty13+"&qty12="+qty12+"&qty11="+qty11+"&qty10="+qty10+"&qty9="+qty9+"&qty8="+qty8+"&qty7="+qty7+"&qty6="+qty6+"&qty5="+qty5+"&qty4="+qty4+"&qty3="+qty3+"&qty2="+qty2+"&qty1="+qty1+"&sku20="+sku20+"&sku19="+sku19+"&sku18="+sku18+"&sku17="+sku17+"&sku16="+sku16+"&sku15="+sku15+"&sku14="+sku14+"&sku13="+sku13+"&sku12="+sku12+"&sku11="+sku11+"&sku10="+sku10+"&sku9="+sku9+"&sku8="+sku8+"&sku7="+sku7+"&sku6="+sku6+"&sku5="+sku5+"&sku4="+sku4+"&sku3="+sku3+"&sku2="+sku2+"&sku1="+sku1+"&notes="+notes+"&size="+size+"&model="+model+"&make="+make+"&repair_num="+repair_num+"&ph_num="+ph_num+"&l_name=" + l_name + "&f_name=" + f_name + "&action=insert";
  

    var request = jQuery.ajax({
      crossDomain: true,
      url: url ,
      method: "GET",
      dataType: "jsonp"


    });
    
    ask_da_q(f_name,l_name,ph_num,'bike');
  }

  // print the returned data
  function ctrlq(e) {
  
	
	$("#re").html(e.result);
	$("#re").css("visibility","visible");
	read_value();
	
  }
  
  

  
  function read_value() {

$("#re").css("visibility","hidden");
   
   document.getElementById("loader").style.visibility = "visible";
    var url = script_url+"?action=read";

$.getJSON(url, function (json) {

    // Set the variables from the results array
   
   
  

        // CREATE DYNAMIC TABLE.
        var table = document.createElement("table");

		

        var header = table.createTHead();
		    var row = header.insertRow(0);  

        var cell1 = row.insertCell(1);
        var cell2 = row.insertCell(2);
        var cell3 = row.insertCell(3);
        var cell4 = row.insertCell(4);
        var cell5 = row.insertCell(5);
        var cell6 = row.insertCell(6);	
        var cell7 = row.insertCell(7);
        var cell8 = row.insertCell(8);	
        var cell9 = row.insertCell(9);
        var cell10 = row.insertCell(10);
        var cell11 = row.insertCell(11);
        var cell12 = row.insertCell(12);
        var cell13 = row.insertCell(13);
        var cell14 = row.insertCell(14);
        var cell15 = row.insertCell(15);
        var cell16 = row.insertCell(16);
        var cell17 = row.insertCell(17);
        var cell18 = row.insertCell(18);
        var cell19 = row.insertCell(19);
        var cell20 = row.insertCell(20);
        var cell21 = row.insertCell(21);
        var cell22 = row.insertCell(22);	
        var cell23 = row.insertCell(23);	
        var cell24 = row.insertCell(24);	
        var cell25 = row.insertCell(25);	
        var cell26 = row.insertCell(26);	
        var cell27 = row.insertCell(27);	
        var cell28 = row.insertCell(28);	
        var cell29 = row.insertCell(29);	
        var cell30 = row.insertCell(30);	
        var cell31 = row.insertCell(31);	
        var cell32 = row.insertCell(32);	
        var cell33 = row.insertCell(33);	
        var cell34 = row.insertCell(34);	
        var cell35 = row.insertCell(35);	
        var cell36 = row.insertCell(36);	
        var cell37 = row.insertCell(37);	
        var cell38 = row.insertCell(38);	
        var cell39 = row.insertCell(39);	
        var cell40 = row.insertCell(40);	
        var cell41 = row.insertCell(41);
        var cell42 = row.insertCell(42);
        var cell43 = row.insertCell(43);
        var cell44 = row.insertCell(44);
        var cell45 = row.insertCell(45);
        var cell46 = row.insertCell(46);
        

		cell1.innerHTML = "<b>f_name</b>";
		cell2.innerHTML = "<b>l_name</b>";
    cell3.innerHTML = "<b>ph_num</b>";
    cell4.innerHTML = "<b>repair_num</b>";
		cell5.innerHTML = "<b>make_model</b>";
		cell6.innerHTML = "<b>notes</b>";
		cell7.innerHTML = "<b>Sku1</b>";
		cell8.innerHTML = "<b>sku2</b>";
		cell9.innerHTML = "<b>sku3</b>";
		cell10.innerHTML = "<b>sku4</b>";
		cell11.innerHTML = "<b>sku5</b>";
		cell12.innerHTML = "<b>sku6</b>";
		cell13.innerHTML = "<b>sku7</b>";
		cell14.innerHTML = "<b>sku8</b>";
		cell15.innerHTML = "<b>sku9</b>";
    cell16.innerHTML = "<b>sku10</b>";
    cell17.innerHTML = "<b>sku11</b>";		
    cell18.innerHTML = "<b>sku12</b>";		
    cell19.innerHTML = "<b>sku13</b>";		
    cell20.innerHTML = "<b>sku14</b>";		
    cell21.innerHTML = "<b>sku15</b>";		
    cell22.innerHTML = "<b>sku16</b>";
    cell23.innerHTML = "<b>sku17</b>";		
		cell24.innerHTML = "<b>sku18</b>";		
    cell25.innerHTML = "<b>sku19</b>";
    cell26.innerHTML = "<b>sku20</b>";		
    cell27.innerHTML = "<b>qty1</b>";		
    cell28.innerHTML = "<b>qty2</b>";		
    cell29.innerHTML = "<b>qty3</b>";		
    cell30.innerHTML = "<b>qty4</b>";		
    cell31.innerHTML = "<b>qty5</b>";		
    cell32.innerHTML = "<b>qty6</b>";		
    cell33.innerHTML = "<b>qty7</b>";		
    cell34.innerHTML = "<b>qty8</b>";		
    cell35.innerHTML = "<b>qty9</b>";		
    cell36.innerHTML = "<b>qty10</b>";		
    cell37.innerHTML = "<b>qty11</b>";		
    cell38.innerHTML = "<b>qty12</b>";		
    cell39.innerHTML = "<b>qty13</b>";		
    cell40.innerHTML = "<b>qty14</b>";		
    cell41.innerHTML = "<b>qty15</b>";
    cell42.innerHTML = "<b>qty16</b>";
    cell43.innerHTML = "<b>qty17</b>";
    cell44.innerHTML = "<b>qty18</b>";
    cell45.innerHTML = "<b>qty19</b>";
    cell46.innerHTML = "<b>qty20</b>";

        
        // ADD JSON DATA TO THE TABLE AS ROWS.
        for (var i = 0; i < json.records.length; i++) {

            tr = table.insertRow(-1);
				var tabCell = tr.insertCell(-1);
                         tabCell.innerHTML = json.records[i].f_name;
	    	tabCell = tr.insertCell(-1);
				        tabCell.innerHTML = json.records[i].l_name;
        tabCell = tr.insertCell(-1);
                		tabCell.innerHTML = json.records[i].ph_num;
        tabCell = tr.insertCell(-1);
				        tabCell.innerHTML = json.records[i].repair_num;
        tabCell = tr.insertCell(-1);
                        tabCell.innerHTML = json.records[i].make_model;
        tabCell = tr.insertCell(-1);
                            tabCell.innerHTML = json.records[i].notes;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku1;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku2;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku3;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku4;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku5;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku6;
        tabCell = tr.insertCell(-1);
                    tabCell.innerHTML = json.records[i].sku7;
        tabCell = tr.insertCell(-1);		
                            tabCell.innerHTML = json.records[i].sku8;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku9;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku10;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku11;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku12;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku13;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku14;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku15;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku16;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku17;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku18;
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku19;
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].sku20;
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty1;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty2;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty3;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty4;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty5;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty6;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty7;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty8;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty9;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty10;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty11;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty12;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty13;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty14;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty15;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty16;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty17;		
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty18;
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty19;
        tabCell = tr.insertCell(-1);		
                        tabCell.innerHTML = json.records[i].qty20;
            }
      

        // FINALLY ADD THE NEWLY CREATED TABLE WITH JSON DATA TO A CONTAINER.
        var divContainer = document.getElementById("showData");
        divContainer.innerHTML = "";
        divContainer.appendChild(table);
		document.getElementById("loader").style.visibility = "hidden";
		$("#re").css("visibility","visible");
    });
	}