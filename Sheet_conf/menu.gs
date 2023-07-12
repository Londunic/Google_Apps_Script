function onOpen(){
  addMenu();
}

function addMenu(){
  var menu = SpreadsheetApp.getUi().createMenu("Forms 📜");
  menu.addItem("Add Opportunity","linkNewOpp");
  menu.addItem("Add Business Downsize","linkNewBDS");
  menu.addToUi();
}

function linkNewOpp(){
  var url = "https://script.google.com/a/macros/globant.com/s/AKfycbyouPN9zXkbj0qosu0ERjzRR0L2YeJ96vNCqUQBUv_Z35jPQ_dEiuhXVe0YC34tY3Nt_g/exec?parametro=form_addNewOpp";
  abrirEnlace(url);
}

function linkNewBDS(){
  var url = "https://script.google.com/a/macros/globant.com/s/AKfycbyouPN9zXkbj0qosu0ERjzRR0L2YeJ96vNCqUQBUv_Z35jPQ_dEiuhXVe0YC34tY3Nt_g/exec?parametro=form_addBDSize";
  abrirEnlace(url);
}

function abrirEnlace(url) {
  var htmlOutput = HtmlService.createHtmlOutput('<script>window.open("' + url + '");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening link...');
}