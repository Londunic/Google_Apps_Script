// Correo Semanal de gráficas

function emailGrafico(){

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Old_data");
  const charts = hoja.getCharts();

  const chartBlobs = new Array(); 
  const emailImages = {};

  var body = "<html><body style='margin: 0; padding: 0;'>";
  body += "<div style='background-color: #F5F6F7; margin: 0 auto; padding: 30px; max-width: 900px;font-family: Arial;'>";
  //Logos
  var logoGlobantLink = "https://drive.google.com/uc?export=view&id=1R7kLBfjaIN132U8YQuR3EKKz5eGv7A2I";
  var logoJJLink = "https://drive.google.com/uc?export=view&id=1oh-1ZCvpeo8kwz4-Tk5f1M3ymYn2uqKh";
  body += "<div style='background-color: #FFFFFF; margin: 0 auto; padding: 20px; overflow: hidden;'>";
  body += '<img src="' + logoGlobantLink + '" style="width: 150px; height: auto;float: left;">';
  body += '<img src="' + logoJJLink + '" style="width: 200px; height: auto;float: right;">';
  body += "</div>";
  //Titulo 1
  body += "<h2 style='background-color: #444444 ;color: white;'><center>Chats</center></h2>";
  //Marco blanco de las tablas
  body += "<div style='background-color: #FFFFFF; margin: 0 auto; padding: 30px; border-radius: 10px;'>";

  charts.forEach(
    function(chart, i){
      chartBlobs[i] = chart.getAs("image/png");
      body += "<p align='center'><img src='cid:chart"+i+"'></p>";
      emailImages["chart"+i] = chartBlobs[i];
    }
  )

  body += "</div>";
  body += "<br></br>";
  //Botón para hoja de excel
  var linkHoja = "https://docs.google.com/spreadsheets/d/1ZZY1kw0_CAPIJWJzmmUFspmQXsUZYfOSI_WN4uyfr1E/edit#gid=734026991";
  body += "<a href='" + linkHoja + "' target='_blank' style='text-decoration:none;'>";
  body += "<button style='border: none;padding: 10px;display: block; margin: 0 auto;background-color: #BFD732;color: black;border-radius: 20px;'>View More</button>";
  body += "</a>";
  body += "</div>";
  body += "</body></html>";
  

  MailApp.sendEmail({
    to: "edwar.londono@globant.com",
    subject: "Charts",
    htmlBody: body,
    inlineImages:emailImages});  

}