function setQty(){
    var select = document.getElementsByClassName('skus');
    var qty = document.getElementsByClassName('qty');
    
    
    for(var i = 0; i<select.length;i++){
        if(select[i].value != ''){
            qty[i].value = '1';
        }
    }
    }
    
    function custyInfo(current,next,div_enabled,next_enabled,next_btn,button_div,button_type){
    
        var to_monkee = setValid(current,button_type);
        
        if (to_monkee == true || button_type == 'back'){
    
            fuck_these_buttons(div_enabled,next_enabled,next_btn,button_div);
            do_your_magic(current,next);
        }
    }
    function test_phone(phone){
    
        phone = phone.replace(/[^\d]/g, '');
        console.log(phone.length);
        return phone.length;
    }
    function phoneBuilder(number){
        var format = "(xxx) xxx-xxxx";
    
        for (var i=0;i<number.length;i++){
            format = format.replace('x', number[i]);
        }
        return format;
    }
    function setValid(current,button_type){
        var poop_time = current.querySelectorAll('.custy');
        
    
        var to_monkee = true;
            for (var i=0;i<poop_time.length;i++){
                
                if (poop_time[i].value == ''){
                    poop_time[i].style.border = 'groove 2.7px rgb(0 255 8 )';
                    poop_time[i].placeholder = 'Please Enter ' + poop_time[i].name;
                    to_monkee = false;
                }
                if (poop_time[i].value == '' && button_type == 'back'){
                    poop_time[i].style.border = '';
                    poop_time[i].placeholder = '';
                }
                if (poop_time[i].value != '' && poop_time[i].name != "Phone Number"){
                    poop_time[i].style.border = '';
                    poop_time[i].placeholder = '';
                }
                if (poop_time[i].name == "Phone Number" && button_type != 'back') {
                    if (test_phone(poop_time[i].value) > 10 || test_phone(poop_time[i].value) < 10){
                        poop_time[i].style.border = 'groove 2.7px rgb(0 255 8 )';
                    poop_time[i].value = '';
                    poop_time[i].placeholder = 'Please Enter ' + poop_time[i].name;
                    to_monkee = false;
                    }
                    if (test_phone(poop_time[i].value) == 10) {
                        poop_time[i].style.border = '';
                        poop_time[i].placeholder = '';
                        var phoneNumba = phoneBuilder(poop_time[i].value.replace(/[^\d]/g, ''));
                        poop_time[i].value = phoneNumba;
                    }
                }
                if (poop_time.value == 'height_select'){
                    poop_time[i].style.border = 'groove 2.7px rgb(0 255 8 )'
                }
    
            }
        return to_monkee;
    }