/* Adapted from https://github.com/NGuernse/SherpaRomeo_GoogleScript which adapted the original code from an 
article by Stephen X. Flynn, Catalina Oyler, and Marsha Miles - http://journal.code4lib.org/articles/7825.

**********************************************************************************************************************
NOTE: Use the Journal Conditions or Restrictions (Pre-Print, Post-Print, or Publshers PDF) to confirm that an article version is allowed in an institutional
      repository (IR). The ArticleVersion returns which version is allowed for self-archiving but does
      not confirm archiving in an IR.
**********************************************************************************************************************/
var sherpaAPIKey = "" // <--- PUT Sherpa/RoMEO API KEY HERE between the quotes

// Add a custom menu to the active spreadsheet
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [];
    // When the user selects "SerpaRoMEO Script" menu, and clicks a functionName, the function is executed
    menuEntries.push({
        name: "Create Column Headers",
        functionName: "createColumnHeaders"
    });
    menuEntries.push({
        name: "Process ISSNs",
        functionName: "processISSNs"
    });
    ss.addMenu("SherpaRoMEO", menuEntries);

}

function processISSNs() {
    plainText(); // Converts ISSN from Automatic to Plain Text
    correctISSNs(); // Fix ISSN numbers
    var sheet = SpreadsheetApp.getActiveSheet();
    var startRow = 2;
    var startColumn = 1;
    var numRows = sheet.getLastRow();
    var numColumns = 11; // number of columns to include from spreadsheet
    var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns)
    var data = dataRange.getValues();
    for (var i = 0; i < data.length; ++i) {
        var column = data[i];
        var ISSN = column[0]; // ISSN column
        if (ISSN != "") { // Process only if ISSN is set
            if (column[1] == "") { // if Journal Title column is empty
                sheet.getRange(startRow + i, 2).setValue(jTitle(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[2] == "") { // if Journal Conditions column is empty
                sheet.getRange(startRow + i, 3).setValue(journalConditions(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[3] == "") { // if Pre-Print Restrictions column is empty
                sheet.getRange(startRow + i, 4).setValue(preRestrictions(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[4] == "") { // if Post-Print Restrictions column is empty
                sheet.getRange(startRow + i, 5).setValue(postRestrictions(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[5] == "") { // if Publishers PDF Restrictions column is empty
                sheet.getRange(startRow + i, 6).setValue(pdfRestrictions(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[6] == "") { // if Copyright Links column is empty
                sheet.getRange(startRow + i, 7).setValue(copyrightLinks(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[7] == "") { // if Sherpa/RoMEO Link column is empty
                sheet.getRange(startRow + i, 8).setValue(sherpaLink(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[8] == "") { // if Embargo column is empty
                sheet.getRange(startRow + i, 9).setValue(embargo(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[9] == "") { // if Article Version column is empty
                sheet.getRange(startRow + i, 10).setValue(articleVersion(ISSN)); // get value from SHERPA/RoMEO
            }
            if (column[10] == "") { // if OA Mandate column is empty
                sheet.getRange(startRow + i, 11).setValue(checkOAmandate(ISSN)); // get value from SHERPA/RoMEO
            }
        }
    }
}

// Check ISSNs for missing digits, etc. and fix if necessary
function correctISSNs() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var startRow = 2;
    var startColumn = 1;
    var numRows = sheet.getLastRow();
    var numColumns = 1;
    var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns)
    var data = dataRange.getValues();
    for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        var ISSN = row[0];
        if (ISSN != "") { // Process only if ISSN is set
            sheet.getRange(startRow + i, 1).setValue(fixISSN(ISSN));
        }
    }

}

// Add column headers to the active spreadsheet
function createColumnHeaders() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    // Freezes the first row
    sheet.setFrozenRows(1);

    // Set the values for headers
    var values = [
        ["ISSN", "Journal Title", "Journal Conditions", "Pre-Print Restrictions", "Post-Print Restrictions", "Publisher's PDF Restrictions", "Copyright Links", "Sherpa/RoMEO Link", "Embargo", "Article Version", "OA Mandate Embargo Imposed"]
    ];

    // Set the range of cells.
    var range = sheet.getRange("A1:K1");

    // Call the setValues method on range and pass in pre-set values
    range.setValues(values);

    // Bold headers
    range.setFontWeight("bold");
}

// Convert ISSN column to Plain Text
function plainText() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var column = sheet.getRange("A2:A");
    column.setNumberFormat('@STRING@');
}

/** Function to fix ISSNs if the starting zeros were deleted
 * @customfunction
 */
function fixISSN(issn) {
    issn = fixissn(issn);
    if (issn == 00000000 || issn == 0000 - 0000 || issn == "") {
        return ("blank ISSN")
    } else if (issn.length > 9) {
        return ("invalid issn")
    } else
        return issn;
}

//These are all the "Helper" functions used in the main functions
function fixissn(issn) {
    Logger.log("Old issn " + issn);

    //convert to string value
    x = issn.toString();

    //if there's a dash in the ISSN, it needs 9 characters instead of 8
    if (x.search("-") > -1) {
        var mis = 9;
    } else if (x.search("-") == -1) {
        var mis = 8;
    }

    //if the ISSN is less than 9 or 8 (mis), add zeros to the beginning
    while (x.length < mis && x.length > 0) {
        //zero added to beginning of ISSN
        x = 0 + x
    }
    Logger.log("fixed to new issn " + x);
    return x;
}

// getXML function taking ISSN and API key
function getXML(issn) {
    // Retrieve XML output from Sherpa API using ISSN input
    issn = fixissn(issn);

    var parameters = {
        method: "get"
    };
    var xmlText = UrlFetchApp.fetch("http://www.sherpa.ac.uk/romeo/api29.php?issn=" + issn + "&versions=all" + "&ak=" + sherpaAPIKey, parameters).getContentText();
    return (xmlText)
}

/** Function to look for "Embargo" in journal's copyright record
 * @customfunction
 */
function embargo(issn) {
    var text = getXML(issn);

    // Check if API limit exceded
    if (checkAPILimit(text)) {
        return ("API Limit Reached")
    }

    // Check ISSN and record validity and whether there is a conditions section
    if (checkISSN(issn, text)) {
        return ("No record or ISSN invalid")
    }

    var embargoPDF = text.search("embargo");
    if (embargoPDF == "-1") {
        embargoPDF = "NO"
    } else if (embargoPDF > 0) {
        embargoPDF = "YES"
    }  
    return embargoPDF;
}

// Check if the ISSN is a journal which specifically forbids uploading to Open Access mandate institutional repositories
function checkOAmandate(issn) {
    var text = getXML(issn);
    // Check to see if the ISSN is invalid
    var failText = text.search("<outcome>failed</outcome>");
    var notfound = text.search("<outcome>notFound</outcome>");
    // Check to see if the ISSN is missing
    if (issn == 00000000 || issn == 0000 - 0000 || issn == "") {
        return ("blank ISSN")
    } else if (failText > 0) {
        return ("ISSN invalid")
    } else if (notfound > -1) {
        return ("not found")
    } else if (failText == -1) {
        return antiOAmandate(text);
    }
}

function permPdfGet(text) {
    var pdfCan = text.search("<pdfarchiving>can</pdfarchiving>");
    var pdfRestricted = text.search("<pdfarchiving>restricted</pdfarchiving>")
    var pdfCannot = text.search("<pdfarchiving>cannot</pdfarchiving>");
    var pdfUnknown = text.search("<pdfarchiving>unknown</pdfarchiving>");
    var pdfUnclear = text.search("<pdfarchiving>unclear</pdfarchiving>");
    var oaPub = text.search("DOAJ says it is an open access journal");
    var noIR = text.search("Cannot post on archival website or institutional repository");
    var noIR2 = text.search("cannot deposit on archival website or institutional repository");

    // If specific negative phrase is found or if <pdfarchiving> is "cannot"
    if (text.search("<pdfarchiving>") == -1) {
        return ("Not Present") // Check if pdfarchiving is even present
    } else if (pdfUnclear > -1) {
        return ("unclear") // Check for unclear policies
    } else if ((pdfCannot > -1) || (noIR > -1) || (noIR2 > -1)) {
        var permText = "No publisher PDF"
    } else if (pdfUnknown > -1) {
        permTxt = "Unknown"
        // If Open Access journal, but not "pdf may be used"
    } else if (oaPub > -1 && pdfCan == "-1") {
        permText = "Open Access journal, check PDF"
    } else if (pdfCan > -1) {
        // If <pdfarchiving> is "can"
        permText = "Publisher's version/PDF may be used"
    } else if (pdfRestricted > -1) {
        // If <pdfarchiving> is "restricted" with an embargo
        permText = "Publisher's version/PDF may be used after an embargo period"
    }  
    return permText;
}

function permFinal(txt) {
    var finalCan = txt.search("<postarchiving>can</postarchiving>");
    var finalPerm = txt.search("<postarchiving>restricted</postarchiving>");
    var finalPermNO = txt.search("<postarchiving>cannot</postarchiving>");
    var finalPermIDK = txt.search("<postarchiving>unknown</postarchiving>");
    var finalUnclear = txt.search("<postarchiving>unclear</postarchiving>");

    if (txt.search("<postarchiving>") == -1) {
        return ("Not Present")
    } // Check if postarchiving is even present
    else if (finalUnclear > -1) {
        return ("unclear")
    } // Check for unclear policies
    else if (finalCan > -1) {
        return ("Final draft allowed")
    } // Check if post-print archiving is allowed
    else if (finalPerm == "-1" && finalPermNO == "-1" && finalPermIDK == "-1") {
        var finalPerm = "Final draft allowed"
    } else if (finalPerm > -1) {
        finalPerm = "Final draft restricted"
    } else if (finalPermNO > -1) {
        finalPerm = "No final draft allowed"
    } else if (finalPermIDK > -1) {
        finalPerm = "status unknown"
    }  
    return finalPerm;
}

// Some publishers like Elsevier single out institutions with Open Access mandates
function antiOAmandate(txt) {
    var antiOA = txt.search("separate agreement between repository and publisher exists");

    if (antiOA > -1) {
        var antiOAstatus = "YES"
    } else if (antiOA == "-1") {
        antiOAstatus = "NO"
    }  
    return antiOAstatus;
}

/** Function to print highest accepted article version
 * @customfunction
 */
function articleVersion(ISSN) {
  
  var text = getXML(ISSN);

    // Check if API limit exceded
    if (checkAPILimit(text)) {
        return ("API Limit Reached")
    }

    // Check ISSN and record validity and whether there is a conditions section
    if (checkISSN(ISSN, text)) {
        return ("No record or ISSN invalid")
    } else {

        var pubPrint = permPdfGet(text);
        var postPrint = permFinal(text);

        if (pubPrint == "Publisher's version/PDF may be used" || pubPrint == "Publisher's version/PDF may be used after an embargo period") {
            return ("Publisher's version may be used")
        } else if (postPrint == "Final draft allowed" || postPrint == "Final draft restricted") {
            return ("Post-print may be used")
        } else if (pubPrint == "unknown" || pubPrint == "unclear" || postPrint == "status unknown" || postPrint == "unclear") {
            return ("unknown/unclear")
        } else if (pubPrint == "Not Present" && postPrint == "Not Present") {
            return ("No record or ISSN invalid")
        } else {
            return ("Neither Publisher's nor Post-print")
        }
      
    }
}

/** Function to print out journal title
 * @customfunction
 */
function jTitle(issn) {

    var xml = getXML(issn);

    // Check if API limit exceeded
    if (checkAPILimit(xml)) {
        return ("API Limit Reached");
    }

    //  Check ISSN and record validity and whether there is a title
    if (checkISSN(issn, xml)) {
        return ("No record or ISSN invalid");
    }

    var document = XmlService.parse(xml);
    var jtitle = document.getRootElement().getChild('journals').getChild('journal').getChild('jtitle').getValue();

    return jtitle;
}

/**
 * Function to print out pre-print restrictions
 * @customfunction
 */

function preRestrictions(issn) {

    var xml = getXML(issn);

    // Check if API limit exceeded
    if (checkAPILimit(xml)) {
        return ("API Limit Reached");
    }

    //  Check ISSN and record validity and whether there is a title
    if (checkISSN(issn, xml)) {
        return ("No record or ISSN invalid");
    }

    var document = XmlService.parse(xml);

    var prerestrictionsText = "";
    try {
        var prerestrictions = document.getRootElement().getChild('publishers').getChild('publisher').getChild('preprints').getChild('prerestrictions').getChildren();

        for (i = 0; i < prerestrictions.length; i++) {
            prerestrictionsText += (i + 1) + '. ' + prerestrictions[i].getValue() + '\n';
        }

    } catch (e) {
        Logger.log(e);
    }
    if (prerestrictionsText == "" || prerestrictionsText == null) {
        prerestrictionsText = "No Restrictions Found";
    } else {
        prerestrictionsText = prerestrictionsText.substring(0, prerestrictionsText.length - 1);
    }

    return prerestrictionsText;
}

/**
 * Function to print out post-print restrictions
 * @customfunction
 */

function postRestrictions(issn) {

    var xml = getXML(issn);

    // Check if API limit exceeded
    if (checkAPILimit(xml)) {
        return ("API Limit Reached");
    }

    //  Check ISSN and record validity and whether there is a title
    if (checkISSN(issn, xml)) {
        return ("No record or ISSN invalid");
    }

    var document = XmlService.parse(xml);

    var postrestrictionsText = "";
    try {
        var postrestrictions = document.getRootElement().getChild('publishers').getChild('publisher').getChild('postprints').getChild('postrestrictions').getChildren();

        for (i = 0; i < postrestrictions.length; i++) {
            postrestrictionsText += (i + 1) + '. ' + postrestrictions[i].getValue()+ '\n';
        }

    } catch (e) {
        Logger.log(e);
    }
    if (postrestrictionsText == "" || postrestrictionsText == null) {
        postrestrictionsText = "No Restrictions Found";
    } else {
        postrestrictionsText = postrestrictionsText.substring(0, postrestrictionsText.length - 1);
    }

    return postrestrictionsText;
}

/**
 * Function to print out publisher's PDF restrictions
 * @customfunction
 */

function pdfRestrictions(issn) {

    var xml = getXML(issn);

    // Check if API limit exceeded
    if (checkAPILimit(xml)) {
        return ("API Limit Reached");
    }

    //  Check ISSN and record validity and whether there is a title
    if (checkISSN(issn, xml)) {
        return ("No record or ISSN invalid");
    }

    var document = XmlService.parse(xml);

    var pdfRestrictionsText = "";
    try {
        var pdfRestrictions = document.getRootElement().getChild('publishers').getChild('publisher').getChild('pdfversion').getChild('pdfrestrictions').getChildren();

        for (i = 0; i < pdfRestrictions.length; i++) {
            pdfRestrictionsText += (i + 1) + '. ' + pdfRestrictions[i].getValue() + '\n';
        }

    } catch (e) {
        Logger.log(e);
    }
    if (pdfRestrictions == "" || pdfRestrictions == null) {
        pdfRestrictionsText = "No Restrictions Found";
    } else {
        pdfRestrictionsText = pdfRestrictionsText.substring(0, pdfRestrictionsText.length - 1);
    }

    return pdfRestrictionsText;
}

/**
 * Function to print out journal conditions
 * @customfunction
 */
function journalConditions(issn) {

    var xml = getXML(issn);

    // Check if API limit exceeded
    if (checkAPILimit(xml)) {
        return ("API Limit Reached");
    }

    //  Check ISSN and record validity and whether there is a title
    if (checkISSN(issn, xml)) {
        return ("No record or ISSN invalid");
    }

    var document = XmlService.parse(xml);

    var conditionsText = "";
    try {
        var conditions = document.getRootElement().getChild('publishers').getChild('publisher').getChild('conditions').getChildren();

        for (i = 0; i < conditions.length; i++) {
            conditionsText += (i + 1) + '. ' + conditions[i].getValue() + '\n';
        }

    } catch (e) {
        Logger.log(e);
    }
    if (conditionsText == "") {
        conditionsText = "No Conditions Found";
    } else {
        conditionsText = conditionsText.substring(0, conditionsText.length - 1);
    }

    return conditionsText;
}

/** Function to print out journal copyright links
 * @customfunction
 */
function copyrightLinks(issn) {
    var xml = getXML(issn);

    // Check if API limit exceeded
    if (checkAPILimit(xml)) {
        return ("API Limit Reached");
    }

    //  Check ISSN and record validity and whether there is a title
    if (checkISSN(issn, xml)) {
        return ("No record or ISSN invalid");
    }

    var document = XmlService.parse(xml);

    var copyrightText = "";
    try {
      var copyright = document.getRootElement().getChild('publishers').getChild('publisher').getChild('copyrightlinks').getChildren();//.getChildren('copyrightlinkurl'); // getChild('copyrightlink') // getChild('copyrightlinkurl'); 
      var x = 1;
        for (i = 0; i < copyright.length; i++) {
          if(copyright[i].getChild('copyrightlinkurl').getValue().toLowerCase().indexOf('http') != -1) {
            copyrightText += x + ". " + copyright[i].getChild('copyrightlinktext').getValue() + ": " + copyright[i].getChild('copyrightlinkurl').getValue() + '\n';
            x += 1;
          }
        }

    } catch (e) {
        Logger.log(e);
    }
    if (copyrightText == "") {
        copyrightText = "No Copyright Links Found";
    } else {
        copyrightText = copyrightText.substring(0, copyrightText.length - 1); // remove last line break
    }
  
    return copyrightText;
}

/** Function to print out Sherpa/RoMEO record link
 * @customfunction
 */
function sherpaLink(issn) {
    issn = fixissn(issn);
    return ("http://www.sherpa.ac.uk/romeo/search.php?issn=" + issn)
}

function checkISSN(issn, text, check) {
    // Set default value for check
    check = check || 0;

    // Check to see if the ISSN is invalid
    var failText = text.search("<outcome>failed</outcome>");
    var notfound = text.search("<outcome>notFound</outcome>");

    // Check if Publisher policies have been checked
    var noPolicyCheck = text.search("<condition>This publisher's policies have not been checked by RoMEO.</condition>");

    // Check to see if the ISSN is missing
    if (issn == 00000000 || issn == 0000 - 0000 || issn == "") {
        return (true)
    } else if (failText > 0) {
        return (true)
    } else if (notfound > -1) {
        return (true)
    } else if (check != 0 && text.search(check) == -1) {
        return (true)
    } else if (noPolicyCheck > -1) {
        return (true)
    } else {
        return (false)
    }

}

// Function to check if api limit is exceeded
function checkAPILimit(text) {
    if (text.search("<outcome>pastfreelimit</outcome>") > -1) {
        return (true)
    } else {
        return (false)
    }
}
