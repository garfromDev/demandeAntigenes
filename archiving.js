/// l'ID du dossier dans lequel des pdf seront archivés
let PDF_FOLDER = "1P9IbI2fZRH6N_S2tVSeu02r9Q4mzFX8H";
/// la colonne dans laquelle insérer le lien
let COL_LINK = 1;
/// la feuille récapitulatif
let RECAP_SHEET = 'Récapitulatif demandes';


/**
 * Archive la demande en pdf après avoir demandé confirmation à l'utilisateur
 * la demande est la feuille pointée par le lien hypertexte de la ligne sélectionné
 * en colonne COL_LINK
 * NOTE : le script lancé par l'utilisateur ne peut pas effacer la feuille protégée, il l'ajoute
 * donc à une liste de feuille à effacer, un script admin passera toute les nuits pour effacer les feuilles
 */
function archiveSelectedLine() {
    let selected = activeSheet().getActiveRange();
    if(selected.getNumRows() > 1){
        alert("Ne sélectionner qu'une ligne à la fois!");
        return
    }
    let ui = SpreadsheetApp.getUi();
    let row = selected.getRow();
    let nom = getSheetNameFromLine(row) || "";
    var response = ui.alert("Confirmer archivage de la ligne ?\nL'opération est irréversible", nom, ui.ButtonSet.YES_NO);
    // Process the user's response.
    if (response == ui.Button.YES){
        archiveLine(row);
    }
  }
  
 
  /**
   * Perform the archiving of the line, handling interaction with user
   * @param {int} line 
   */
function archiveLine(line){
    toast(`Archivage de la ligne ${line}...`);
    let sheet = getSheetFromLine(line);
    if(!sheet){
        alert(`Impossible d'archiver la ligne ${line}, déjà archivée ou données erronées`);
        return;
    }
    toast("Création du pdf...");
    let pdf = exportToPdf(sheet);
    if(!pdf){
        alert(`La génération du fichier pdf ${sheet.getName()} pour la ligne ${line} a échouée`);
    }
   
    markToBeDeleted(sheet, line, pdf.getUrl());
    toast("Génération du pdf terminée avec succès, la page sera effacée automatiquement dans la nuit...");
}


function exportToPdf(sheet) {
    // 1 create a new sheet
    let newSpreadSheet = SpreadsheetApp.create('___TEMP___SPREADSHEET_FOR_PDF');
    // 2 copy sheet to new spreadsheet, delete sheet1 which is first sheet
    let newSheet = sheet.copyTo(activeSpreadSheet());
    let data = sheet.getDataRange();
    data.copyTo(newSheet.getRange('A1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    data.copyTo(newSheet.getRange('A1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    data.copyTo(newSheet.getRange('A1'), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);    
    newSheet.copyTo(newSpreadSheet);
    newSpreadSheet.deleteSheet(newSpreadSheet.getSheets()[0]);
    activeSpreadSheet().deleteSheet(newSheet);
    // 3 create pdf from new spreadsheet, rename it like old sheet
    let pdfFolder = DriveApp.getFolderById(PDF_FOLDER);
    let pdfData = newSpreadSheet.getAs('application/pdf');
    let pdf = DriveApp.createFile(pdfData);
    pdfFolder.addFile(pdf);
    pdf.setName(sheet.getName());
    // 6 delete temp sheet
    let tempFile = DriveApp.getFileById(newSpreadSheet.getId());
    DriveApp.removeFile(tempFile);

    return pdf;
}


/**
 * @param {int} line 
 * @returns {Sheet|bool}  the sheet which is referenced by the hyperlink of the line
 * false if no link found
 */
function getSheetFromLine(line) {
    let formula = getFormulaFromLine(line);
    let findId = /(?:\#gid\=)(.*)(?:\";)/g; // l'id est entre #gid= et le "; 
    let potentialId = findId.exec(formula);
    let id = potentialId && potentialId.length > 0 ? potentialId[1]: false || false;
    return id? getSheetById(id): false;
}


/**
 * @param {int} line 
 * @returns {String} name of the sheet that is referenced by the hypertext link
 * it is assumed that the display value is the name of the sheet
 * return false if no value found for any reason
 */
function getSheetNameFromLine(line){
    let formula = getFormulaFromLine(line) || "";
    let findId = /(?:\";\s*\")(.*)(?:\"\))/g; // l'id est entre #gid= et le "; 
    let potentialId = findId.exec(formula);
    let id = potentialId && potentialId.length > 0 ? potentialId[1]: false || false;
    return id;    
}


/**
 * Retourne la formule de la cellule en colonne COL_LINK de la feuille
 * active pour la ligne line
 * @param {int} line 
 */
function getFormulaFromLine(line){
    let cell = activeSheet().getRange(line, COL_LINK);
    return cell.getFormula();
}


/**
 * Note la feuille comme étant à effacer en onscrivant son nom
 * à la fin de la liste de la feuille __TO_BE_DELETED__
 * @param {Sheet} sheet 
 */
function markToBeDeleted(sheet, line, link) {
    let sh = getSheet('__TO_BE_DELETED__');
    let l = getLastRowForColumn(sh.getRange('A:A'), true) + 1;
    sh.getRange(l, 1).setValue(sheet.getName());
    sh.getRange(l, 2).setValue(line);
    sh.getRange(l, 3).setValue(link);
}


/**
 * ****************************************************************
 * This script must be run with admin right to override protections
 * it will delete all sheets whose name are listed in column A of 
 *  masqued sheet __TO_BE_DELETED__
 * and transform line of recap into value only greyed line, with
 * link to the pdf file
 */
function deleteMarkedSheets() {
    let sh = getSheet('__TO_BE_DELETED__');
    let names = sh.getRange('A:C').getValues();
    let recap = getSheet(RECAP_SHEET);
    var l = getLastRowForColumn(sh.getRange('A:A'))
    while(l > 0){
        let name = names[l-1][0];
        let line = names[l-1][1];
        let pdfLink = names[l-1][2];
         // 4 copy value from old line
        copyValueOnly(line, 'grey');
        // 5 insert link to pdf
        linkCell = recap.getRange(line, COL_LINK);
        addHyperlinkToCell(linkCell, pdfLink);
        // delete  old sheet
        try{
            activeSpreadSheet().deleteSheet(getSheet(name));
        }catch{
            // we do nothing because if the user click many time archiving,
            // the line will be duplicated, and the sheet may have been already deleted
        }
        sh.getRange(l--, 1).setValue("");
    }
}
