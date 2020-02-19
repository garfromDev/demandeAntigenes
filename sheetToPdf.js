
// thanks to valentin https://www.nexus-creation.com/2019/06/25/google-apps-script-envoyer-une-feuille-par-mail-en-piece-jointe-pdf/
//Convert spreadsheet to PDF file.
/**
 * 
 * @param {*} id : id de la spreadsheet
 * @param {*} index gid de la feuille
 * @param {*} name le nom du fichier a créer
 * @returns {Blob} un blob de type pdf
 */
function sheetToPDF(id, index, name)
{
SpreadsheetApp.flush();
//make OAuth connection
var token = ScriptApp.getOAuthToken();
//get request
var request = {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
};

//define the params URL to fetch
var params = 
  'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&size=A4'                       // paper size legal / letter / A4
  + '&portrait=false'                    // orientation, false for landscape
  + '&fitw=true&source=labnol'           // fit to page width, false for actual size
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&gid=' + index;

//fetching file url
var blob = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/"+id+"/export?"+params, request);
blob = blob.getBlob().setName(name);
//return file
return blob;
}
