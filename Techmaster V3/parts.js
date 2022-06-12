function addPart(type){
    var part_input = document.getElementById(type+'_part_input');
    var elements1 = document.getElementsByClassName(type+'_part');
    var elements2 = document.getElementsByClassName(type+'_part_div');
    var elements3 = document.getElementsByClassName('qty');
    
    for(var i=0; i<10;i++){
        if(elements1[i].value == "" && part_input.value != "" && part_input.value.length > 0 && part_input.value.length <=12){
            elements2[i].style.display = 'flex';
            elements1[i].value = part_input.value; 
            elements3[i+20].value = '1';
            part_input.value = '';
            break;
        }
        if (elements1[i].value != "" && part_input.value != "" && i == 9 && part_input.value.length > 0 && part_input.value.length <=12){
            alert("You can only insert 10 skus");
            part_input.value = '';
            break;
        }
        if(part_input.value == "" || part_input.value != "" && part_input.value.length < 0 || part_input.value.length >12){
            alert("Please enter a correct value");
            part_input.value = '';
            break;
        }
    }
}