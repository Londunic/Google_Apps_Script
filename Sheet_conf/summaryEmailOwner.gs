// Correo Semanal a Operation and Tech Owners

function summaryEmailOwner(){
  var opps = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("(New) Opportunity Tracking");
  var dataOpss = opps.getRange("A2:W").getValues();

  var ownersData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Data Sets");
  var owners = ownersData.getRange("D2:D").getValues();
  var ownersEmails = ownersData.getRange("D2:E").getValues();

  //Se extraen los owners únicos
  var auxOwner = [];
  for(var i = 0; i < owners.length; i++){
    auxOwner.push(owners[i][0]);
  }
  var ownersUnicos = [...new Set(auxOwner)];
  // Se eliminan los valores nulls
  ownersUnicos = ownersUnicos.filter(Boolean);


  //Se extraen los correos de los owners
  var emails = [];
  for(var i = 0; i < ownersUnicos.length; i++){
    for(var k = 0; k < ownersEmails.length; k++){
      if(ownersUnicos[i] == ownersEmails[k][0]){
        emails.push(ownersEmails[k][1]);
        break;
      }
    }
  }

  //Se crean los correos de los Operation Owners
  for(var i = 0; i < ownersUnicos.length; i++){
    
    //Se filtran las opps asociadas al owner y las que se encuentren OnGoing
    var ownerData = [];

    for(var k = 0; k < dataOpss.length; k++){
      if(dataOpss[k][3] == ownersUnicos[i] && (dataOpss[k][12] == "20% - Pipeline/Lead" || dataOpss[k][12] == "40% - Qualified Lead" || dataOpss[k][12] == "60% - Quoted/SOW" || dataOpss[k][12] == "80% - Shortlisted") && dataOpss[k][13] != ""){
        ownerData.push(dataOpss[k]);
      }
    }
    
    //Verifico si el owner tiene opps OnGoing para Mandar el correo
    if(ownerData.length > 0){
      //Ordeno la data por Target Date, de menor a mayor.
      ownerDataOrdenada = ordenarOpps(ownerData);
      //Creo y envio correo a owner
      var body = bodyOwnerEmail("Operation Owner",ownersUnicos[i],ownerDataOrdenada)
      var subject = "Correo Prueba";
      var recipients = "edwar.londono@globant.com";
      // var recipients = emails[i];
      MailApp.sendEmail({to: recipients, subject: subject,htmlBody: body});
    }
  }

  //Se crean los correos de los Tech Owners
  for(var i = 0; i < ownersUnicos.length; i++){
    
    //Se filtran las opps asociadas al owner y las que se encuentren OnGoing
    var ownerData = [];

    for(var k = 0; k < dataOpss.length; k++){
      if(dataOpss[k][4] == ownersUnicos[i] && (dataOpss[k][12] == "20% - Pipeline/Lead" || dataOpss[k][12] == "40% - Qualified Lead" || dataOpss[k][12] == "60% - Quoted/SOW" || dataOpss[k][12] == "80% - Shortlisted")&& dataOpss[k][13] != ""){
        ownerData.push(dataOpss[k]);
      }
    }
    
    //Verifico si el owner tiene opps OnGoing para Mandar el correo
    if(ownerData.length > 0){
      //Ordeno la data por Target Date, de menor a mayor.
      ownerDataOrdenada = ordenarOpps(ownerData);
      //Creo y envio correo a owner
      var body = bodyOwnerEmail("Tech Owner",ownersUnicos[i],ownerDataOrdenada)
      var subject = "Correo Prueba";
      var recipients = "edwar.londono@globant.com";
      // var recipients = emails[i];
      MailApp.sendEmail({to: recipients, subject: subject,htmlBody: body});
    }
  } 
}

function ordenarOpps(dataOpss){
  
  dataOrdenada = [];

  while(dataOpss.length>0){
    var indexMayor = -1;
    var bigTarDate = "";
    for(var k = 0; k < dataOpss.length; k++){
      if(dataOpss[k][13] > bigTarDate){
        bigTarDate = dataOpss[k][13]
        indexMayor = k;
      }
    }
    dataOrdenada.unshift(dataOpss[indexMayor]);
    dataOpss.splice(indexMayor, 1);
  }

  return dataOrdenada;
}

function bodyOwnerEmail(position,ownerName,ownerData){

  var body = "<html><body style='margin: 0; padding: 0;'>";
  body += "<div style='background-color: #F5F6F7; margin: 0 auto; padding: 30px; max-width: 1100px;font-family: Arial;'>";

  //Logos
  var logoGlobantLink = "https://drive.google.com/uc?export=view&id=1R7kLBfjaIN132U8YQuR3EKKz5eGv7A2I";
  var logoJJLink = "https://drive.google.com/uc?export=view&id=1oh-1ZCvpeo8kwz4-Tk5f1M3ymYn2uqKh";
  body += "<div style='background-color: #FFFFFF; margin: 0 auto; padding: 20px; overflow: hidden;'>";
  body += '<img src="' + logoGlobantLink + '" style="width: 150px; height: auto;float: left;">';
  body += '<img src="' + logoJJLink + '" style="width: 200px; height: auto;float: right;">';
  body += "</div>";

  //Titulo 1
  body += "<h2 style='background-color: #444444 ;color: white;'><center>My Ongoing Opportunities As "+ position +"</center></h2>";
  
  //Titulo 2: Nombre el Owner
  body += "<h3 style='background-color: #BFD732 ;color: white;'><center>"+ ownerName +"</center></h3>";
  //Marco blanco de las tablas
  body += "<div style='background-color: #FFFFFF; margin: 0 auto; padding: 30px; border-radius: 10px;'>";
  body += "<hr></hr>";

  //Se muestran las distintas opps.
  for(var k = 0; k < ownerData.length; k++){

    body += "<table align='center' border='0' cellpadding='0' cellspacing='0' width='100%'>";

    body += "<tr>";
    body += "<th style='width: 8%;text-align: center;'>"+ "N° Opp" +"</th>";
    body += "<th style='width: 14%;text-align: center;'>"+ "Op. Name" +"</th>";
    body += "<th style='width: 14%;text-align: center;'>"+ "Stakeholder" +"</th>";
    body += "<th style='width: 12%;text-align: center;'>"+ "Stage" +"</th>";
    body += "<th style='width: 12%;text-align: center;'>"+ "Monthly RR" +"</th>";
    body += "<th style='width: 12%;text-align: center;'>"+ "Weighted Amt" +"</th>";
    body += "<th style='width: 14%;text-align: center;'>"+ "Full Amt" +"</th>";
    body += "<th style='width: 14%;text-align: center;'>"+ "Target Date" +"</th>";
    body += "</tr>";

    //Formato Pesos a Monthly RunRate - Total Amount - Full Amount
    var valorFormateado1 = ownerData[k][17].toLocaleString('en-US', { style: 'currency', currency: 'USD' });
    var valorFormateado2 = ownerData[k][20].toLocaleString('en-US', { style: 'currency', currency: 'USD' });
    var valorFormateado3 = ownerData[k][22].toLocaleString('en-US', { style: 'currency', currency: 'USD' });

    //Formato a Target Date
    var fecha = new Date(ownerData[k][13]);
    fecha.setDate(fecha.getDate());
    fechaNew1 = Utilities.formatDate(fecha, "GMT", "MM/dd/yyyy");

    body += "<tr>";
    body += "<td style='width: 8%;text-align: center;'>"+ ownerData[k][0] +"</td>";
    body += "<td style='width: 14%;text-align: center;'>"+ ownerData[k][7] +"</td>";
    body += "<td style='width: 14%;text-align: center;'>"+ ownerData[k][5] +"</td>";
    body += "<td style='width: 12%;text-align: center;'>"+ ownerData[k][12] +"</td>";
    body += "<td style='width: 12%;text-align: center;'>"+ valorFormateado1 +"</td>";
    body += "<td style='width: 12%;text-align: center;'>"+ valorFormateado2 +"</td>";
    body += "<td style='width: 14%;text-align: center;'>"+ valorFormateado3 +"</td>";
    body += "<td style='width: 14%;text-align: center;'>"+ fechaNew1 +"</td>";
    body += "</tr>";
    
    body += "</table>";
    body += "<hr></hr>";
  }
  body += "</div>";
  body += "<br></br>";
  //Botón para hoja de excel
  var linkHoja = "https://docs.google.com/spreadsheets/d/1ZZY1kw0_CAPIJWJzmmUFspmQXsUZYfOSI_WN4uyfr1E/edit#gid=734026991";
  body += "<a href='" + linkHoja + "' target='_blank' style='text-decoration:none;'>";
  body += "<button style='border: none;padding: 10px;display: block; margin: 0 auto;background-color: #BFD732;color: black;border-radius: 20px;'>View More</button>";
  body += "</a>";
  body += "</div>";
  body += "</body></html>";
  
  return body;
}