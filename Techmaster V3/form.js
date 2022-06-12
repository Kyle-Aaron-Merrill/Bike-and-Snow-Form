document.body.addEventListener('keypress', (e) => {
    var type = document.querySelectorAll('.button_dec');
    type.forEach(element => {
        if(element.style.display == 'flex') type = element.id.split('_')[0];
    });
    var element = document.getElementsByName('next')
    element.forEach(btn => {
        if(btn.id.split('_')[0] == type) element = btn;
    });
    if(e.key == 'Enter') button_mover(element);
});

function button_mover(that){
    var div_enabled = true;
    var next_enabled = true;

    var type = that.id.split("_")[0];
    var button_type = that.name;
    
    var button_div = document.getElementById(type + "_button_div");
    var btns = document.querySelectorAll("." + type + "_button");

    var next_btn = btns[1];

    var elements = document.querySelectorAll(".element");

    for(var i=0;i<elements.length;i++){
        var current = elements[i];
        if (current.style.display == 'flex') {
            if (button_type == 'next') {
                for(var n=i;n<elements.length;n++){
                    var next = elements[n+1];
                    
                    selectChecker(next.id.split('_')[1],type);
                    serviceChecker(current.id.split('_')[0],current.id.split('_')[1]);
                    next_enabled = notesChecker(next.id.split('_')[1]);
                    
                    var isFalse = next.classList.contains("false");

                    if (isFalse != true){
                        custyInfo(current,next,div_enabled,next_enabled,next_btn,button_div,button_type);
                        break
                    }
                }
            }
            if (button_type == 'back') {
                for(var n=i;n>=0;n--){
                    var next = elements[n-1];
                    var obj = typeChecker(current,next,elements);
                    next = obj[0];
                    div_enabled = obj[1];

                    selectChecker(next.id.split('_')[1],type);
                    var isFalse = next.classList.contains("false");

                    if (isFalse != true){  
                        custyInfo(current,next,div_enabled,next_enabled,next_btn,button_div,button_type);
                        break
                    }
                }
            }
        }
    }
}

function selectChecker(id,type){
    var align = document.getElementById(type +'_alignAll');
    if (id == 'service' || id == 'search') {
        setTimeout(function(){
            align.style.margin = 'auto';
        }, 0500);
    }
    if (id != 'service' && id != 'search') {
        setTimeout(function(){
            //align.style.margin.remove('auto')
            align.style.margin = '20% auto';
        }, 0500);
    }
}
function fuck_these_buttons(div_enabled,next_enabled,next,button_div){
    if (div_enabled == false) {
        button_div.classList.remove('fade_in');
        button_div.classList.add('fade_out');
        setTimeout(function(){
            button_div.style.display = 'none';
        }, 0500);
    }
    if (next_enabled == false) {
        button_div.classList.remove('fade_in');
        button_div.classList.add('fade_out');
        setTimeout(function(){
            next.style.display = 'none';
        }, 0500);
        setTimeout(function(){    
            button_div.classList.add('fade_in');
        }, 0600);
    }
    if (div_enabled == true && next_enabled == true) {
        button_div.classList.remove('fade_in');
        button_div.classList.add('fade_out');
        setTimeout(function(){
            next.style.display = 'flex';
            button_div.style.display = 'flex';
        }, 0500);
        setTimeout(function(){    
            button_div.classList.add('fade_in');
        }, 0600);
    }
}
function typeChecker(element,next,array,){
    var div_enabled = true;
    var obj = [next,div_enabled];
    var id = element.id.split('_')[0];
    if(element.id.split('_')[0] == 'bike' && element.id.split('_')[1] == 'custy' || next.id.split('_')[0] == 'type')
    {
        obj[0] = array[0];
        obj[1] = false;
        document.getElementById(id+"_tech").style.display = 'none';
    }
    
    return obj;
}   

function serviceChecker(type,id){
    if (id == "service") {
        run_dem_skus_ya_heard(type);
    }
}
function notesChecker(id){
    var enable_next = true;
    if (id == "notes") {
        enable_next = false;
    }
    return enable_next;
}

function do_your_magic(current,next){
    current.classList.remove('fade_in');
    current.classList.add('fade_out');
    setTimeout(function(){
        current.style.display = 'none';
        next.style.display = 'flex';
    }, 0500);
    setTimeout(function(){    
        next.classList.add('fade_in');
    }, 0600);
    console.log(current.id);
}
function enable_buttons(type){
    var button_div = document.getElementById(type + "_button_div");
    var div_enabled = true;
    var next_enabled = true;
    var btns = document.querySelectorAll("." + type + "_button");
    var next = btns[1];

    fuck_these_buttons(div_enabled,next_enabled,next,button_div);
}
function enable_form(type){
    var form = document.getElementById(type+"_tech");
    
    setTimeout(function(){
        form.style.display = 'flex';
    }, 0500);
    
}