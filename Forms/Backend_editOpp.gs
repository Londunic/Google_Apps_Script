function infoOpp(value){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("(New) Opportunity Tracking");
  var nOpps = hoja.getRange("A2:A").getValues();

  var fila = 0;
  var encontrado = false;
  for(var i = 0; i < nOpps.length; i++){
    if(nOpps[i][0] == value){
      fila = (i+2);
      encontrado = true;
      break;
    }
  }

  var valores = [];
  if(encontrado){
    var data = hoja.getRange(fila, 1, 1, 23).getValues();
    var cadena="";
    for(var i = 0; i < data[0].length; i++){
      if(i == 11 || i == 13 || i == 14 || i == 15 || i == 18){
        cadena += formatoFecha(data[0][i]);
      }
      else if(i == 17){
        cadena += parseInt(data[0][i]).toLocaleString('en-US', { style: 'currency', currency: 'USD' });
      }
      else{
        cadena +=data[0][i];
      }
      
      if(i != (data[0].length-1)){
        cadena += "||"  
      }
    }
    valores.push(cadena)
  }

  return valores
}


function formatoFecha(fecha){
  if(fecha == ""){
    return "";
  }
  else{
    var date = new Date(fecha);
    date.setDate(date.getDate() + 1);
    var month = date.getMonth() + 1;
    var year = date.getFullYear();

    var mesText = "";

    if(month==1){
      mesText="January";
    }else if(month==2){
      mesText="February";
    }else if(month==3){
      mesText="March";
    }else if(month==4){
      mesText="April";
    }else if(month==5){
      mesText="May";
    }else if(month==6){
      mesText="June";
    }else if(month==7){
      mesText="July";
    }else if(month==8){
      mesText="August";
    }else if(month==9){
      mesText="September";
    }else if(month==10){
      mesText="October";
    }else if(month==11){
      mesText="November";
    }else if(month==12){
      mesText="December";
    }
    mesText = mesText+" "+year;
    return mesText;
  }
}

function editOpp(valorNuevo,nOpp,nombreColumna){

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("(New) Opportunity Tracking");
  var nOpps = hoja.getRange("A2:A").getValues();

  if(nombreColumna == "Business Vertical"){
    var col = 2;
  }
  else if(nombreColumna == "Project name"){
    var col = 3;
  }
  else if(nombreColumna == "Operation Owner *"){
    var col = 4;
  }
  else if(nombreColumna == "Tech Owner"){
    var col = 5;
  }
  else if(nombreColumna == "J&J Stakeholder"){
    var col = 6;
  }
  else if(nombreColumna == "Technology | Tool | Platform"){
    var col = 7;
  }
  else if(nombreColumna == "Opportunity Name"){
    var col = 8;
  }
  else if(nombreColumna == "Opportunity Description"){
    var col = 9;
  }
  else if(nombreColumna == "Last Updates"){
    var col = 10;
  }
  else if(nombreColumna == "Next Actions"){
    var col = 11;
  }
  else if(nombreColumna == "Stage *"){
    var col = 13;
  }
  else if(nombreColumna == "Target Date *"){
    var col = 14;
  }
  else if(nombreColumna == "Opportunity BornDate"){
    var col = 15;
  }
  else if(nombreColumna == "Openings (# of positions)"){
    var col = 17;
  }
  else if(nombreColumna == "Monthly RunRate (All) *"){
    var col = 18;
  }
  else if(nombreColumna == "Opportunity End Date *"){
    var col = 19;
  }
  var fila = 0;
  var encontrado = false;
  for(var i = 0; i < nOpps.length; i++){
    if(nOpps[i][0] == nOpp){
      fila = (i+2);
      encontrado = true;
      break;
    }
  }

  if(encontrado){
    hoja.getRange(fila, col).setValue(valorNuevo);
  }

}



function prueba(){
  Logger.log(infoOpp("100101"))
}



function prueba2(){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("(New) Opportunity Tracking");
  var data = hoja.getRange(2, 1, 1, 23).getValues();
  var fechaCambiada = formatoFecha(data[0][14]);
  Logger.log(fechaCambiada)
}