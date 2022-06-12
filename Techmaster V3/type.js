function getForm(that){
    var elements = document.querySelectorAll('.element');
    var type = that.id.split("_");
    type = type[0];
    var current = document.getElementById('type_sel');
    var next = document.getElementById(type+"_custy_info");

    enable_form(type);
    do_your_magic(current,next);
    enable_buttons(type);
}
function return_to_screen(type){
    var form = document.getElementById(type+'_tech');
    form.reset();
    
    var current = document.getElementById(type+'_notes_div')
    var next = document.getElementById('type_sel');

    var div_enabled = false;
    var next_enabled = false;
    var button_div = document.getElementById(type + "_button_div");


    fuck_these_buttons(div_enabled,next_enabled,next,button_div);
    do_your_magic(current,next);
    document.getElementById('ski_tech').style.display = 'none';
    document.getElementById('bike_tech').style.display = 'none';
}
function ask_da_q(f_name,l_name,ph_num,type){
    var first = f_name;
    var last = l_name;
    var phone = ph_num;
    var name = f_name + " " + l_name;
    var txt = name + "'s ticket was submitted" + '\n' + "click OK to submit another ticket for " + name; 
    
    if (confirm(txt)) {
            return_to_screen(type)
            document.getElementById(type+"_f_name").value = first;
            document.getElementById(type+"_l_name").value = last;
            document.getElementById(type+"_ph_num").value = phone;
    } 
    else {
        return_to_screen(type);
    }
        
}