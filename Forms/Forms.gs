function doGet(e) {
  var parametro = e.parameter.parametro;

  if (parametro === "form_addNewOpp") {
    return form_Add_New_Opp();
  } else if (parametro === "form_addBDSize") {
    return form_Business_Down_Size();
  } else if (parametro === "form_editOPP") {
    return form_Edit_Opp();
  }
}

function form_Add_New_Opp() {
  return HtmlService.createHtmlOutputFromFile('form_addNewOpp');
}

function form_Business_Down_Size() {
  return HtmlService.createHtmlOutputFromFile('form_addBDSize');
}

function form_Edit_Opp() {
  return HtmlService.createHtmlOutputFromFile('form_editOPP');
}