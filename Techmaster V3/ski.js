var ski_url = "https://script.google.com/macros/s/AKfycbzowCaNR7VgDkTBi6GAmfmypzDWpbGTfvq9djfJUqJ77TvttXCFN_JpebG3euc4H3S4bQ/exec";
    //This Script is for the google web app/api
    // Make an AJAX call to Google Script
        

    function insert_ski() {
        
        var f_name= $("#ski_f_name").val();
        var l_name= $("#ski_l_name").val();
        var ph_num= $("#ski_ph_num").val();
        var blank = $("#blank").val();
        var blank2 = $("#blank2").val();
        var ski_snow= $('input[name=ski_snow]:checked').val();
        var weight= $("#weight").val();
        var height= $("#height").val();
        var skier_type= $('input[name=skier_type]:checked').val();
        var stance= $('input[name=stance]:checked').val();
        var age= $("#age").val();
        var make= $("#ski_make").val();
        var mdl= $("#ski_model").val();
        var size= $("#ski_size").val();
        var boot_make= $("#boot_make").val();
        var boot_mdl= $("#boot_model").val();
        var bindings= $("#bindings").val();
        var bootLength= $("#boot_length").val();
        var notes= $("#ski_notes").val();
        var sku_1= $("#ski_sku1").val();
        var sku_2= $("#ski_sku2").val();
        var sku_3= $("#ski_sku3").val();
        var sku_4= $("#ski_sku4").val();
        var sku_5= $("#ski_sku5").val();
        var sku_6= $("#ski_sku6").val();
        var sku_7= $("#ski_sku7").val();
        var sku_8= $("#ski_sku8").val();
        var sku_9= $("#ski_sku9").val();
        var sku_10= $("#ski_sku10").val();
        var qty_1= $("#ski_qty_1").val();
        var qty_2= $("#ski_qty_2").val();
        var qty_3= $("#ski_qty_3").val();
        var qty_4= $("#ski_qty_4").val();
        var qty_5= $("#ski_qty_5").val();
        var qty_6= $("#ski_qty_6").val();
        var qty_7= $("#ski_qty_7").val();
        var qty_8= $("#ski_qty_8").val();
        var qty_9= $("#ski_qty_9").val();
        var qty_10= $("#ski_qty_10").val();


        var url = ski_url+"?callback=ctrlq&age="+age+"&bootLength="+bootLength+"&bindings="+bindings+"&boot_mdl="+boot_mdl+"&boot_make="+boot_make+"&size="+size+"&skier_type="+skier_type+"&height="+height+"&weight="+weight+"&stance="+stance+"&ski_snow="+ski_snow+"&notes="+notes+"&qty_10="+qty_10+"&qty_9="+qty_9+"&qty_8="+qty_8+"&qty_7="+qty_7+"&qty_6="+qty_6+"&qty_5="+qty_5+"&qty_4="+qty_4+"&qty_3="+qty_3+"&qty_2="+qty_2+"&qty_1="+qty_1+"&sku_10="+sku_10+"&sku_9="+sku_9+"&sku_8="+sku_8+"&sku_7="+sku_7+"&sku_6="+sku_6+"&sku_5="+sku_5+"&sku_4="+sku_4+"&sku_3="+sku_3+"&sku_2="+sku_2+"&sku_1="+sku_1+"&mdl="+mdl+"&make="+make+"&blank2="+blank2+"&blank="+blank+"&ph_num="+ph_num+"&l_name=" + l_name + "&f_name=" + f_name + "&action=insert";
    

        var request = jQuery.ajax({
        crossDomain: true,
        url: url ,
        method: "GET",
        dataType: "jsonp"
        });
        
        ask_da_q(f_name,l_name,ph_num,'ski');
        
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
    var url = ski_url+"?action=read";

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
                
            

          cell1.innerHTML = "<b>f_name</b>";
          cell2.innerHTML = "<b>l_name</b>";
          cell3.innerHTML = "<b>ph_num</b>";
          cell4.innerHTML = "<b>blank<b>";
          cell5.innerHTML = "<b>blank<b>";
          cell6.innerHTML = "<b>make</b>";
          cell7.innerHTML = "<b>mdl</b>";

          cell8.innerHTML = "<b>sku_1</b>";
          cell9.innerHTML = "<b>sku_2</b>";
          cell10.innerHTML = "<b>sku_3</b>";
          cell11.innerHTML = "<b>sku_4</b>";
          cell12.innerHTML = "<b>sku_5</b>";
          cell13.innerHTML = "<b>sku_6<b>"; 
          cell14.innerHTML = "<b>sku_7<b>";
          cell15.innerHTML = "<b>sku_8<b>";
          cell16.innerHTML = "<b>sku_9<b>";
          cell17.innerHTML = "<b>sku_10<b>";

          cell18.innerHTML = "<b>qty_1<b>";
          cell19.innerHTML = "<b>qty_2<b>";  
          cell20.innerHTML = "<b>qty_3<b>";  
          cell21.innerHTML = "<b>qty_4<b>";
          cell22.innerHTML = "<b>qty_5<b>";
          cell23.innerHTML = "<b>qty_6<b>";
          cell24.innerHTML = "<b>qty_7<b>";  
          cell25.innerHTML = "<b>qty_8<b>";  
          cell26.innerHTML = "<b>qty_9<b>";
          cell27.innerHTML = "<b>qty_10<b>";       
                    
                             

          cell28.innerHTML = "<b>notes</b>";
          cell29.innerHTML = "<b>date_in</b>";
          cell30.innerHTML = "<b>date_out</b>";
          cell31.innerHTML = "<b>ski_snow</b>";
          cell32.innerHTML = "<b>stance</b>"
          cell33.innerHTML = "<b>weight</b>";
          cell34.innerHTML = "<b>height</b>";
          cell35.innerHTML = "<b>skier_type</b>";
          cell36.innerHTML = "<b>size</b>";       
          cell37.innerHTML = "<b>boot_make</b>";      
          cell38.innerHTML = "<b>boot_mdl</b>";       
          cell39.innerHTML = "<b>bindings</b>";
          cell40.innerHTML = "<b>bootLength<b>";
          cell41.innerHTML = "<b>age<b>";                 
            

            
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
                            tabCell.innerHTML = json.records[i].blank;                
            tabCell = tr.insertCell(-1);
                            tabCell.innerHTML = json.records[i].blank2;                
            tabCell = tr.insertCell(-1);
                            tabCell.innerHTML = json.records[i].make;
            tabCell = tr.insertCell(-1);
                            tabCell.innerHTML = json.records[i].mdl;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_1;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_2;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_3;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_4;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_5;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_6;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_7;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_8;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_9;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].sku_10;
            tabCell = tr.insertCell(-1);             
                            tabCell.innerHTML = json.records[i].qty_1;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_2;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_3;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_4;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_5;
            tabCell = tr.insertCell(-1);             
                            tabCell.innerHTML = json.records[i].qty_6;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_7;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_8;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_9;      
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].qty_10;
            tabCell = tr.insertCell(-1);
                            tabCell.innerHTML = json.records[i].notes;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].date_in;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].date_out;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].ski_snow;
            tabCell = tr.insertCell(-1);               
                            tabCell.innerHTML = json.records[i].stance;               
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].weight;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].height;
            tabCell = tr.insertCell(-1);
                            tabCell.innerHTML = json.records[i].size;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].boot_make;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].boot_mdl;                       
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].bindings;  
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].bootLength;
            tabCell = tr.insertCell(-1);        
                            tabCell.innerHTML = json.records[i].age;
                
                
        
                }
        

            // FINALLY ADD THE NEWLY CREATED TABLE WITH JSON DATA TO A CONTAINER.
            var divContainer = document.getElementById("showData");
            divContainer.innerHTML = "";
            divContainer.appendChild(table);
            document.getElementById("loader").style.visibility = "hidden";
            $("#re").css("visibility","visible");
        });
}