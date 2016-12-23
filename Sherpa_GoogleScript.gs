/************************** COMMENTS and REMARKS ***************************************************
The following code has been copied and edited from the original code in the article by
Stephen X. Flynn, Catalina Oyler, and Marsha Miles - http://journal.code4lib.org/articles/7825
Thank you!
***************************        Corrections      ************************************************
I found an error in the permFinal() and pubPDF() functions which returned a false-positive result albeit 
insufficient Sherpa/RoMEO data to confirm a post-print or publisher's version is allowed for archival 
purposes. Corrections have been made and should alleviate this issue.
**************************         Additions        ************************************************
The following additions were made to the code:
1) A function to get all the conditions for a journal's record in Sherpa/RoMEO - journalCondition()
2) A function to get all copyright policies for a journal's record in Sherpa/RoMEO - copyrightLinks()
3) A function to get the journal's Sherpa/RoMEO hyperlink to the record - sherpaLink()
4) Combining the permFinal() and pubPDFGet() functions into one function articleVersion - articleVersion()
5) Added a checkISSN() function so that the code is not repeated in each function.
6) Added a checkAPILimit() function so that useres are notified when non-api key use limits are reached.
 
*************************       Instructions (copied from original code) ********************************
Watch a quick installation video at http://youtu.be/ZMyKVHM5nOc
(A) Register for an API key at http://www.sherpa.ac.uk/romeo/apiregistry.php then insert the key below in the 
"Sherpa API key" section of the code. Without an API key, you will be limited to 500 requests per day, per IP address.
 
(1) Perform an affiliation search in a large bibliographic database, such as Scopus or Web of Knowledge. 
Include the ISSN metadata in the database export, since the function depends on the ISSN to work. Export this search to a .csv file.
(2) Import the .csv into a Google Spreadsheet with at a minimum, columns for the following: ISSN, Article Version, Embargo, Conditions, Copyright links, sherpa link. 
Other details like Journal title, Article title, and Author are optional.
(3) CRITICAL: Select the ISSN column, select Format -> Number -> Plain Text. Or else the script won't work!
(4) Go to Script -> Script Editor
(5) Copy and paste all the following into script editor window, and save.
(6) Call functions described below in their corresponding columns: articleVersion([issn column cell]), embargo([issn column cell],
journalConditions([issn column cell]), copyrightLinks([issn column cell]), sherpaLink([issn column cell]). 
For example, =articleVersion(19352735) will lookup the ISSN 1935-2735 (PLoS Neglected Tropical Diseases) in Sherpa/Romeo, 
and result in the text "Publisher's version may be used".
 
TIP: Go to Format -> Conditional formatting... to set up color codes to help visualize your column outputs.
TIP: Beware of excessive use imposed by Google. Currently you are limited to 20,000-50,000 URL lookups per day. 
     See https://docs.google.com/macros/dashboard for UrlFetch specifically.
TIP: To avoid excessive usage, after running a set of ISSNs, copy and paste over the discovered values by going to Edit -> Paste special -> Paste values only
**********************************************************************************************************************
NOTE: Make sure to double check the conditions to confirm that an article version is allowed in an institutional
      repository. The articleVersion function returns which version is allowed for self-archiving but does
      not confirm archiving in an IR. Using the journalConditions function will help confirm this. Look for a typical
      statement that says "Article is allowed to be archived in an author's institutional repository" or some
      similar statement.
**********************************************************************************************************************/      


//Function: use this function if your system deleted the starting zeros of your ISSN numbers. DSpace 1.6 does this on //a metadata export.
function fixISSNdoc(issn) {
  issn = fixissn(issn);
  if (issn == 00000000 || issn == 0000-0000 || issn == "")
  { return ("blank ISSN")
  } else if (issn.length > 9) {
    return ("invalid issn") } else
      return issn;
}
 
//test function for logging purposes
function testSherp() {
   var result = pubpdf("2041-8205");
   Logger.log("test result is '" + result + "'");
}
 
//These are all the "Helper" functions used in the main functions
function fixissn(issn){
  Logger.log("Old issn " + issn);
  
  //convert to string value
  x = issn.toString();
  
  //if there's a dash in the ISSN, it need 9 characters instead of 8
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

//getXML function taking issn and api key
function getXML(issn){
    // retrieves XML output from Sherpa API using issn input
  issn = fixissn(issn);
 
//Google scripts will timeout if you run the script too often. These sleep actions are designed to space out your commands. Comment out the next 6 lines if you're testing and only running small batches (less than 10 at a time).
    //var randnumber = Math.random()*5000;
    //Utilities.sleep(randnumber);
    //Utilities.sleep(randnumber);
    //Utilities.sleep(randnumber);
    //Utilities.sleep(randnumber);
    //Utilities.sleep(randnumber);
  
  //register for a key at http://www.sherpa.ac.uk/romeo/apiregistry.php
  //enter API key between quotes for sherpaAPIKey
  var parameters = {method : "get"};
  var sherpaAPIKey = ""               
  var xmlText = UrlFetchApp.fetch("http://www.sherpa.ac.uk/romeo/api29.php?issn=" + issn + "&versions=all" + "&ak=" + sherpaAPIKey,parameters).getContentText();
  return(xmlText)
  
}//end getXML function
 
//Function: look for the word "Embargo" in the journal's copyright record
function embargo(issn){
  var text = getXML(issn);
  
  // Check if API limit exceded
  if (checkAPILimit(text)){ return("API Limit Reached") }
  
  // Check issn and record validity and whether there is a conditions section
  if(checkISSN(issn,text)) { return("No record or ISSN invalid") }
  
  var embargoPDF=text.search("embargo");
  if (embargoPDF == "-1") {
    embargoPDF="no"
  } else if (embargoPDF > 0 ) {
    embargoPDF = "Embargo"
  } return embargoPDF;
}
 
//checks if the ISSN is a journal which specifically forbids uploading to Open Access mandate institutional repositories
function checkOAmandate(issn){
 var text = getXML(issn);
   // checks to see if the issn is invalid
  var failText=text.search("<outcome>failed</outcome>");
  var notfound=text.search("<outcome>notFound</outcome>");
  // checks to see if the issn is missing
  if (issn == 00000000 || issn == 0000-0000 || issn == "") { return ("blank ISSN")
  } else if (failText > 0){
    return ("ISSN invalid") }
  else if (notfound > -1){
    return ("not found") }
  else if (failText == -1) {  
  return antiOAmandate(text);
  }
}
 
function permPdfGet(text) {
  var pdfCan=text.search("<pdfarchiving>can</pdfarchiving>");
  var pdfRestricted=text.search("<pdfarchiving>restricted</pdfarchiving>")
  var pdfCannot=text.search("<pdfarchiving>cannot</pdfarchiving>");
  var pdfUnknown=text.search("<pdfarchiving>unknown</pdfarchiving>");
  var pdfUnclear=text.search("<pdfarchiving>unclear</pdfarchiving>");
  var oaPub=text.search("DOAJ says it is an open access journal");
  var noIR=text.search("Cannot post on archival website or institutional repository");
  var noIR2=text.search("cannot deposit on archival website or institutional repository");
 
  //if we find specific negative phrase or if <pdfarchiving> is "cannot"
  if (text.search("<pdfarchiving>") == -1) { return("Not Present") //Check if pdfarchiving is even present
  } else if (pdfUnclear > -1) { return("unclear") //Check for unclear policies
  } else if((pdfCannot > -1) || (noIR > -1) || (noIR2 > -1)) {
    var permText = "No publisher PDF"
    } else if (pdfUnknown > -1) {
      permTxt = "Unknown"
  // if its an open access journal, but not "pdf may be used"
  } else if (oaPub > -1 && pdfCan == "-1") {
    permText = "Open access journal, check PDF"
  }  else if (pdfCan > -1) {
  // if <pdfarchiving> is "can"
    permText = "Publisher's version/PDF may be used"
  }  else if (pdfRestricted > -1) {
  // if <pdfarchiving> is "restricted" with an embargo
    permText= "Publisher's version/PDF may be used after an embargo period" }
  return permText;
}
 
function permFinal(txt) {
  var finalCan=txt.search("<postarchiving>can</postarchiving>");
  var finalPerm=txt.search("<postarchiving>restricted</postarchiving>");
  var finalPermNO=txt.search("<postarchiving>cannot</postarchiving>");
  var finalPermIDK=txt.search("<postarchiving>unknown</postarchiving>");
  var finalUnclear=txt.search("<postarchiving>unclear</postarchiving>");
 
  if(txt.search("<postarchiving>") == -1) {return("Not Present")} //Check if postarchiving is even present
  else if(finalUnclear > -1) {return("unclear")}//Check for unclear policies
  else if (finalCan > -1) { return ("Final draft allowed") } //Check if post-print archiving is allowed
  else if(finalPerm == "-1" && finalPermNO == "-1" && finalPermIDK == "-1"){
    var finalPerm = "Final draft allowed"
    }else if (finalPerm > -1) {
      finalPerm = "Final draft restricted"
    }else if (finalPermNO > -1) {
      finalPerm = "NO final draft allowed"
    } else if (finalPermIDK > -1) {
      finalPerm = "status unknown" }
     return finalPerm;
}
 
//some publishers like Elsevier single out institutions with open access mandates
function antiOAmandate(txt) {
 var antiOA=txt.search("separate agreement between repository and publisher exists");
 
  if (antiOA > -1) {
    var antiOAstatus = "Anti OA mandate"
    }  else if (antiOA == "-1") {
      antiOAstatus = "all good!" }
  return antiOAstatus;
}

/********  Original Code Above ******** Additions Below *************************************/

// Function to print highest accepted article version
function articleVersion(issn) {
 var text = getXML(issn);
  
 // Check if API limit exceded
 if (checkAPILimit(text)){ return("API Limit Reached") }
  
 // Check issn and record validity and whether there is a conditions section
 if(checkISSN(issn,text)) { return("No record or ISSN invalid") } else {
  
    var pubPrint = permPdfGet(text);
    var postPrint = permFinal(text);
    
    if (pubPrint == "Publisher's version/PDF may be used" || pubPrint == "Publisher's version/PDF may be used after an embargo period") { 
      return("Publisher's version may be used")
    } else if ( postPrint == "Final draft allowed" || postPrint == "Final draft restricted" ) {
      return("Post-print may be used")
    } else if (pubPrint == "unknown" || pubPrint == "unclear" || postPrint == "status unknown" || postPrint == "unclear") {return("unknown/unclear")
    } else if (pubPrint == "Not Present" && postPrint == "Not Present") { return("No record or ISSN invalid") }
   else { return("Neither Publisher's nor Post-print") }                                                                                                                      
    
  }
} //END OF articleVersion function

// Function to print out journal conditions
function journalConditions(issn) {
  var text = getXML(issn);
  
  // Check if API limit exceded
  if (checkAPILimit(text)){ return("API Limit Reached") }
  
  // Check issn and record validity and whether there is a conditions section
  if(checkISSN(issn,text,"<conditions>")) { return("No record, ISSN invalid, or no copyright links") }
  
  var conditions = text.slice(text.search("<conditions>")+12,text.search("</conditions>")).trim();
  var condlist = ""; // code if you want a list: condlist = [];
  
  for (var c = 1; conditions.length > 0; c++) {
    
    //Extract Condition
    var tempcond = conditions.slice(conditions.search("<condition>")+11,conditions.search("</condition>")).replace(/&lt;([^.]*?)&gt;/gi, "");
    
    //Append to conditions list
    condlist = condlist.concat(c + ". " + tempcond + "\n");
    
    //Update conditions text
    conditions = conditions.slice(conditions.search("</condition>")+12).trim();
    
  } // end of for loop
  
  if(condlist.length == 0) { return("No Conditions.") } else {return(condlist) }
} // END OF journalConditions function

// Function to print out journal copyright links
function copyrightLinks(issn) {
  var text = getXML(issn);
  
  // Check if API limit exceded
  if (checkAPILimit(text)){ return("API Limit Reached") }
  
  // Check issn and record validity and whether there is a copyright section
  if(checkISSN(issn,text,"<copyrightlinks>")) { return("No record, ISSN invalid, or no copyright links") }
  
  var copyrights = text.slice(text.search("<copyrightlinks>")+16,text.search("</copyrightlinks>")).trim();
  var copylist = ""; // code if you want a list: copylist = [];

  
  for (var c = 1; copyrights.length > 0; c++) {
    
    //Extract Copyright URL Text
    var temptext = copyrights.slice(copyrights.search("<copyrightlinktext>")+19,copyrights.search("</copyrightlinktext>"))
    
    //Extract Copyright URL
    var tempurl = copyrights.slice(copyrights.search("<copyrightlinkurl>")+18,copyrights.search("</copyrightlinkurl>"))
    
    //Append to List
    copylist = copylist.concat(c + ". " + temptext + ": " + tempurl + "\n");
    
    //Slice to new set of copyrightlinks
    copyrights = copyrights.slice(copyrights.search("</copyrightlink>")+16);
  } // end of for loop
  
  if(copylist.length == 0) { return("No copyright links.") } else {return(copylist) }
} // END of copyrightLinks function


function sherpaLink(issn) {
  issn = fixissn(issn);
  return("http://www.sherpa.ac.uk/romeo/search.php?issn="+issn.slice(0,4)+"-"+issn.slice(4,8))
} // End of sherpaLink function

function checkISSN(issn,text,check) {
  //Sets default value for check
  check = check || 0;
  
  // checks to see if the issn is invalid
  var failText = text.search("<outcome>failed</outcome>");
  var notfound = text.search("<outcome>notFound</outcome>");
  //check if Publisher policies have been checked.
  var noPolicyCheck = text.search("<condition>This publisher's policies have not been checked by RoMEO.</condition>");
  
  // checks to see if the issn is missing
  if (issn == 00000000 || issn == 0000-0000 || issn == "") { return (true)
  } else if (failText > 0){
    return (true) }
  else if (notfound > -1){
    return (true) }
  else if (check != 0 && text.search(check) == -1) {return(true)}
  else if (noPolicyCheck > -1){
    return(true)
  } else { return(false) }
  
} // end of checkISSN function

// Function to check if api limit is exceeded
function checkAPILimit(text) {
  if(text.search("<outcome>pastfreelimit</outcome>") > -1) { return(true) } else { return(false) }
}// end of checkAPILimit function

