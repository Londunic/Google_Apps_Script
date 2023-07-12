function valuesRangoBDSize(rango){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Data Sets");
  var opciones = hoja.getRange(rango).getValues();
  var valores = opciones.map(function(row) {
    return row[0];
  });
  valores = valores.filter(Boolean);
  return valores;
}

function valuesRangoBDSize2(rango){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stakeholders Map");
  var opciones = hoja.getRange(rango).getValues();
  var valores = opciones.map(function(row) {
    return row[0];
  });
  valores = valores.filter(Boolean);
  return valores;
}

function valuesFechaRangoBDSize(rango) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Data Sets");
  var opciones = hoja.getRange(rango).getValues();
  var valores = opciones.map(function(row) {
    return row[0];
  });
  valores = valores.filter(Boolean);

  var values = [];
  for(var i = 0; i < valores.length; i++){
    var date = new Date(valores[i]);
    date.setDate(date.getDate() + 1);
    var day = date.getDate();
    var month = date.getMonth() + 1;
    var year = date.getFullYear();

    // Formatea la fecha resultante como 'dd-mm-yyyy'
    var formattedDate = month+"/"+day+"/"+ year;

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
    var aux = [mesText,formattedDate];
    values.push(aux);
  }

  return values;
}

function addBDSize(values) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Business Downsize");
  var columna = hoja.getRange("A2:A").getValues();

  //Verifico el id más alto en la hoja
  var indiceMax = 0;
  for(var i = 0; i < columna.length; i++){
    if( parseInt(columna[i][0]) >indiceMax){
      indiceMax=parseInt(columna[i][0]);
    }
  }
  //Creo el índice para el nuevo BDSize
  var id = indiceMax+1;
  var linea= 0;
  
  //Recorro toda la columna A para identificar la primera celda vacia
  for(var i = 0; i < columna.length; i++){

    //Si la celda donde estoy parado es vacia, entonces imprimo
    if(columna[i][0] == ""){

      hoja.getRange(i+2,1).setValue(id);
      for(var k = 0; k < values.length; k++){
        hoja.getRange(i+2,k+2).setValue(values[k]);  
      }
      linea=(i+2);
      break;
    }    
  }
}