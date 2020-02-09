

//fonction appelée par le bouton
function EnvoiMailRespPharma(e) {
   var DEMANDEUR_CELL = "E3";
  var shObj = activeSheet();
  var demandeur = shObj.getRange(DEMANDEUR_CELL).getValue();
  sendEmail("aline.fromont@ceva.com, laurent.drouet@ceva.com, mathilde.hocquard@ceva.com, alain.schrumpf@ceva.com, "+ demandeur,"Nouvelle demande d'antigène",7);
}

//fonction appelée par le bouton
function EnvoiMailProduction(e) {
   var DEMANDEUR_CELL = "E3";
  var shObj = activeSheet();
  var demandeur = shObj.getRange(DEMANDEUR_CELL).getValue();
  sendEmail("aline.fromont@ceva.com, philippe.grel@ceva.com,graziella.bourdet@ceva.com, virginie.fleurie@ceva.com, "+ demandeur,"Nouvelle demande d'antigène",8);
}

//fonction appelée par le bouton
function EnvoiMailDemandeur(e) {
  var DEMANDEUR_CELL = "E3";
  var shObj = activeSheet();
  var demandeur = shObj.getRange(DEMANDEUR_CELL).getValue();
  sendEmail("aline.fromont@ceva.com, clementine.mottais@ceva.com, elodie.gontier@ceva.com, philippe.grel@ceva.com, graziella.bourdet@ceva.com, virginie.fleurie@ceva.com, "+ demandeur,"Nouvelle demande d'antigène saisie et enregistrée",9);
}

/**
Envoi un email depuis le compte de l'utilisateur courant
to : l'adresse (ou les adresses séparés par des virgules) de destination
subject : le sujet du mail
fromCol : le no de la colonne dans laquelle on trouve le contenu du message
*/
function sendEmail(to, subject, fromCol)
{
  //1 on récupère le nO de ligne de la cellule sélectionnée
  l = 46;
  // 2 le contenu du mail est dans la cellule de la même ligne, en colonne fromCol
  contenu = SpreadsheetApp.getActiveSheet().getRange(l, fromCol).getValue();
  
// 3 Display a dialog box with a title, message, and "Yes" and "No" buttons. The
// user can also close the dialog by clicking the close button in its title bar.
var ui = SpreadsheetApp.getUi();
var response = ui.alert('Confirmer envoi email ?', contenu, ui.ButtonSet.YES_NO);

// Process the user's response.
if (response == ui.Button.YES)
  MailApp.sendEmail(to, subject, contenu);
else if (response == ui.Button.NO)
  return;// on arrête tout
else
  Logger.log('The user clicked the close button in the dialog\'s title bar.');
return; // on arrête tout
}  


//fonction pour récupérer l'url d'une feuille
function urlFeuille() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();  
  urlFeuille = spreadsheet.getFormUrl();

} 


/** Custom function  
* utilisation : `= NOM_FEUILLE()`
* @return {String} Le nom de la feuille courante
*/
function NOM_FEUILLE() {
  return activeSheet().getName();
}



/** Custom function  
* utilisation : `= URL_FEUILLE()`
* @return {String} L'URL de la feuille courante
*/
function URL_FEUILLE() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + activeSheet().getSheetId().toString();
}
  

function testexportpdf() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
};