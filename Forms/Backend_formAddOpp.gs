function valuesRango(rango) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Data Sets");
  var opciones = hoja.getRange(rango).getValues();
  var valores = opciones.map(function(row) {
    return row[0];
  });
  valores = valores.filter(Boolean);
  return valores;
}

function valuesFechaRango(rango) {
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

function addNewOpp1(values){
  //  Se ingresa la info en la planilla

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("(New) Opportunity Tracking");
  var columna = hoja.getRange("A2:A").getValues();

  //Verifico el # de opp más alto en la hoja
  var indiceMax = 0;
  for(var i = 0; i < columna.length; i++){
    if( parseInt(columna[i][0]) >indiceMax){
      indiceMax=parseInt(columna[i][0]);
    }
  }
  //Creo el índice para la nueva oportunidad
  var nOPP = indiceMax+1;
  var linea= 0;

  //Last Update
  var lUpdate = "This opportunity was added";
  //Last Update Date
  var fecha = new Date();
  var lUpdateDate = (fecha.getMonth()+1) + "/" + fecha.getDate() + "/" + fecha.getFullYear();
  //Won Date
  var wonDate = "";
  if(values[9] == "100% - Won"){
    wonDate = lUpdateDate;
  }
  
  //Recorro toda la columna A para identificar la primera celda vacia
  for(var i = 0; i < columna.length; i++){

    //Si la celda donde estoy parado es vacia, entonces imprimo
    if(columna[i][0] == ""){

      hoja.getRange(i+2,1).setValue(nOPP);
      hoja.getRange(i+2,10).setValue(lUpdate);
      hoja.getRange(i+2,11).setValue(values[8]);
      hoja.getRange(i+2,12).setValue(lUpdateDate);
      hoja.getRange(i+2,16).setValue(wonDate);
      hoja.getRange(i+2,17).setValue(values[12]);
      hoja.getRange(i+2,18).setValue(values[13]);
      hoja.getRange(i+2,19).setValue(values[14]);

      for(var k = 0; k < 8; k++){
        hoja.getRange(i+2,2+k).setValue(values[k]);  
      }
      for(var k = 9; k < 12; k++){
        hoja.getRange(i+2,4+k).setValue(values[k]);  
      }
      linea=(i+2);
      break;
    }    
  }

  //Meses entre Target Date y End Date
  var months = hoja.getRange(linea,20).getValue();
  //Weighted Amount
  var wAmt = hoja.getRange(linea,21).getValue();
  //Full Amount
  var fAmount = hoja.getRange(linea,23).getValue();

  //  Se crea y se envia el correo

  //Defino las variables que voy a utilizar para enviar el correo
  var opO= values[2];  //Operation Owner
  var teO= values[3];  //Tech Owner
  var bV= values[0];   //Business Vertical
  var sH= values[4];   //StakeHolder
  var oNm= values[6];  //OppName

  var hoja2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Data Sets");
  var users = hoja2.getRange("D2:D").getValues();
  var mails = hoja2.getRange("E2:E").getValues();
  var mailOPO = "";
  var mailTEO = "";

  //Se extrae los correos de Op. Owner y Tech Owner
  var stop1=false;
  var stop2=false;
  for(var i = 0; i < users.length; i++){
    //Extraigo el correo del Operation Owner
    if(users[i][0] == opO ){
      mailOPO = mails[i];
      stop1=true;
    }

    //Extraigo el correo del Tech Owner
    if(users[i][0] == teO ){
      mailTEO = mails[i];
      stop2=true;
    }

    if(stop1==true && stop2==true){
      break;
    }
  }

  //Creo el cuerpo del correo
  var body = "<html><body style='margin: 0; padding: 0;'>";
  body += "<div style='background-color: #F5F6F7; margin: 0 auto; padding: 30px; max-width: 600px;font-family: Arial;'>";

  //Logos
  var logoGlobantLink = "https://drive.google.com/uc?export=view&id=1R7kLBfjaIN132U8YQuR3EKKz5eGv7A2I";
  var logoJJLink = "https://drive.google.com/uc?export=view&id=1oh-1ZCvpeo8kwz4-Tk5f1M3ymYn2uqKh";
  body += "<div style='background-color: #FFFFFF; margin: 0 auto; padding: 20px; overflow: hidden;'>";
  body += '<img src="' + logoGlobantLink + '" style="width: 150px; height: auto;float: left;">';
  body += '<img src="' + logoJJLink + '" style="width: 200px; height: auto;float: right;">';
  body += "</div>";
  //Titulo 1
  body += "<h2 style='background-color: #444444 ;color: white;'><center>A NEW OPPORTUNITY HAS BEEN ADDED</center></h2>";
  //Marco blanco de la tabla
  body += "<div style='background-color: #FFFFFF; margin: 0 auto; padding: 30px;border-radius: 10px;'>";
  //Tabla
  body += "<table align='center' border='0' cellpadding='0' cellspacing='0' width='100%'>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Opp Number</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ nOPP +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Business Vertical</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[0] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Project</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[1] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Operation Owner</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[2] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Tech Owner</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[3] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>J&J Stakeholder</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[4] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Technology|Tool|Platform</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[5] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Opportunity Name</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[6] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Opportunity Description</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[7] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Next Actions</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[8] +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Stage</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[9] +"</td>";
  body += "</tr>";

  //Formato fecha Target Date
  var fecha = new Date(values[10]);
  fecha.setDate(fecha.getDate());
  fechaNew1 = Utilities.formatDate(fecha, "GMT", "MM/dd/yyyy");

  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Target Date</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ fechaNew1 +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Openings(# of positions)</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ values[12] +"</td>";
  body += "</tr>";

  //Formato Pesos a Monthly RunRate
  var valorFormateado1 = parseInt(values[13]).toLocaleString('en-US', { style: 'currency', currency: 'USD' });

  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Monthly RunRate (All)</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ valorFormateado1 +"</td>";
  body += "</tr>";
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Duration (months)</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ months +"</td>";
  body += "</tr>";

  //Formato Pesos a Weighted amount
  var valorFormateado2 = wAmt.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Weighted Amount</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ valorFormateado2 +"</td>";
  body += "</tr>";

  //Formato Pesos a Full amount
  var valorFormateado3 = fAmount.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  body += "<tr>";
  body += "<th style='width: 40%;text-align:right;'>Full Amount</th>";
  body += "<td style='width: 60%;text-align:center;'>"+ valorFormateado3 +"</td>";
  body += "</tr>";
  
  body += "</table>";
  body += "</div>";
  body += "<br></br>";

  //Botón para hoja de excel
  var linkHoja = "https://docs.google.com/spreadsheets/d/19MhlraOMxR6fgne7Y7qZcRgh3Bfoa39aFo9gf4GDz0s/edit#gid=734026991";
  body += "<a href='" + linkHoja + "' target='_blank' style='text-decoration:none;'>";
  body += "<button style='border: none;padding: 10px;display: block; margin: 0 auto;background-color: #BFD732;color: black;border-radius: 20px;'>View More</button>";
  body += "</a>";
  body += "</div>";
  body += "</body></html>";
  
  var subject = nOPP+" / "+bV+" / "+sH+" / "+oNm;
  var recipients = mailOPO+","+mailTEO;
  MailApp.sendEmail({to: recipients, subject: subject,htmlBody: body});
}