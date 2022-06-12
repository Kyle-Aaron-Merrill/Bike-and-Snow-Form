

function run_dem_skus_ya_heard(type){
    var sku1= $('#'+type+'_sku1').val();
    var sku2= $('#'+type+'_sku2').val();
    var sku3= $('#'+type+'_sku3').val();
    var sku4= $('#'+type+'_sku4').val();
    var sku5= $('#'+type+'_sku5').val();
    var sku6= $('#'+type+'_sku6').val();
    var sku7= $('#'+type+'_sku7').val();
    var sku8= $('#'+type+'_sku8').val();
    var sku9= $('#'+type+'_sku9').val();
    var sku10= $('#'+type+'_sku10').val();

    var arr = [sku1,sku2,sku3,sku4,sku5,sku6,sku7,sku8,sku9,sku10];
    var snowElements = document.querySelectorAll('.snow');
    var skiElements = document.querySelectorAll('.ski');
    var ski = false;
    var snow = false;

    ski = ski_skus(arr);
    snow = snow_skus(arr);

    snowElements.forEach(element => {
        element.classList.remove('false');
    });
    skiElements.forEach(element => {
        element.classList.remove('false');
    });

    if (ski == true && snow == false){
        snowElements.forEach(element => {
            element.classList.add('false');
        });
    }
    else if (ski == false && snow == true){
        skiElements.forEach(element => {
            element.classList.add('false');
        });
        
    }
    else if (ski == true && snow == true){
        snowElements.forEach(element => {
            element.classList.remove('false');
        });
        skiElements.forEach(element => {
            element.classList.remove('false');
        });
    }
    else if(ski == false && snow == false){
        snowElements.forEach(element => {
            element.classList.add('false');
        });
        skiElements.forEach(element => {
            element.classList.add('false');
        });
    }
}
function ski_skus(array){
    var to_monkey = false;
    array.forEach(element => {
        if (element == '4234102' || element == '4234103' || element == '4234101' || element == '4234106' || element == '4234108'){  to_monkey = true;} 
    });    
    return to_monkey;
}
function snow_skus(array){
    var to_monkey = false;
    array.forEach(element => {
        if (element == '4234111' || element == '4234105' || element == '4234104'){ to_monkey = true;}
    });
    return to_monkey;
}
