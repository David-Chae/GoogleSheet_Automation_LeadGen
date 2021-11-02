let activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
let source = activeSpreadSheet.getSheets()[0];
let destination = activeSpreadSheet.getSheets()[1];
let lastRowIndex = source.getLastRow();
let numOfRecords = source.getLastRow() - 1;


function onOpen(){
  let uiBuilder = SpreadsheetApp.getUi();
  let menu = uiBuilder.createMenu('Prepare Hubspot Import Menu');
  menu.addItem('Run All Preps' , 'sigmaAll').addToUi();
  menu.addItem('Add Checked Columns','addCheckColumns').addToUi();
  menu.addItem('Modify Imported Table' , 'modifyImportedExcel').addToUi();
  menu.addItem('Fill with Default Values' , 'fillWithDefaultValues').addToUi();
  menu.addItem('Fetch company domain names with company names' , 'fetchAndFillCompanyUrl').addToUi();
  menu.addItem('Decide and select Deal Industry Category' , 'decideDealIndustryCategory').addToUi();
  menu.addItem('Decide and select Deal Category' , 'showHTTPResponse').addToUi();
}


function showHTTPResponse(){
  //Let's try the function to output https://www.google.com/search?q=site%3Awww.linkedin.com+accenture
  /*
  let google = "https://www.google.com/search?q=";
  let constraint_linkedin = "site:www.linkedin.com"
  let company_name = "accenture"
  let searchUrl = google + encodeURIComponent(constraint_linkedin+" "+company_name);
  console.log(google + encodeURIComponent(constraint_linkedin+" "+company_name));
  let response = UrlFetchApp.fetch(searchUrl);
  //console.log(response.getAllHeaders());
  */
  

  let wiki = "https://en.wikipedia.org/wiki/";
  let linkedin = "https://www.linkedin.com/company/";
  let company_name = "Avanade"; 
  let searchUrl = linkedin + encodeURIComponent(" "+company_name);
  let response = UrlFetchApp.fetch(searchUrl);

  let html = HtmlService.createHtmlOutput("'" + response.getContentText() + "'");
  SpreadsheetApp.getUi().showModalDialog(html, "Find company industry");
  Utilities.sleep(10000);

  /*
  for(let i = 0; i < companyNames.length; i++){
    console.log(companyNames[i]);
  }
  */
  //Let's get company names. 
  //destination.getRange(2, getHeaderColumnContaining("Company Name", destination), numOfRecords);


/*
  //for(let i = 0; i < companyNames.length; i++){
    let company_name = companyNames[0];
    let searchUrl = encodeURIComponent(wiki+" "+company_name);
    let response = UrlFetchApp.fetch(searchUrl);
    let html = HtmlService.createHtmlOutput("'" + response.getContentText() + "'");
    SpreadsheetApp.getUi().showModalDialog(html, "Find company industry");
    Utilities.sleep(10000);
  //}
*/

  //I want to get the first search result from google using UrlFetchApp. So, I can cut the first block of search result into the html page in the custom dialog box.



}



function decideDealIndustryCategory(){

  let categories = ["IT", "IT Security", "HR", "Health and Safety", "CFO", "Supply Chain", "Marketing", "Energy"];

  for(let i = 2; i <= lastRowIndex; i++){
    let jobTitle = destination.getRange(i, getHeaderColumnContaining("Job Title", destination)).getValue();
    let html = returnHtmlwith(i);
    SpreadsheetApp.getUi().showModalDialog(html, jobTitle);
    Utilities.sleep(10000);
  }

}


function returnHtmlwith(i){

  let html = HtmlService.createHtmlOutput(
     '<!DOCTYPE html>'+
      '<html>'+
        '<head>'+
          '<base target="_top">'+
          '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">'+
        '</head>'+
        '<body>'+
          '<div>'+
            '<table>'+
              '<tr><td>IT </td>               <td><input type="radio" name="category" value="IT" checked></td></tr>'+
              '<tr><td>IT Security </td>      <td><input type="radio" name="category" value="IT Security"></td></tr>'+
              '<tr><td>HR </td>               <td><input type="radio" name="category" value="HR"></td></tr>'+
              '<tr><td>Health and Safety </td><td><input type="radio" name="category" value="Health and Safety"></td></tr>'+
              '<tr><td>Supply Chain </td>     <td><input type="radio" name="category" value="Supply Chain"></td></tr>'+
              '<tr><td>Marketing </td>        <td><input type="radio" name="category" value="Marketing"></td></tr>'+
              '<tr><td>Energy </td>           <td><input type="radio" name="category" value="Energy"></td></tr>'+       
            '</table>'+
            '<input type="button" value="Submit" class="action" onclick="form_data()" >'+
            '<input type="button" value="Close" onclick="google.script.host.close()" />'+
          '</div>'+
          '<script src="//ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>'+
          '<script>'+
            'function form_data(){'+
              'let value = $("input[name=category]:checked").val();'+
              'google.script.run.withSuccessHandler().returnIndustryCategory(value' + ',' + i + ');'+
              'closeIt()'+
            '};'+
            'function closeIt(){'+
              'google.script.host.close()'+
            '};'+
          '</script>'+
        '</body>'+
      '</html>'
  );

  return html;
}


function returnIndustryCategory(values , i){
  let dealIndustryCategoryCell = destination.getRange(i, getHeaderColumnContaining("Deal Industry Category", destination));
  dealIndustryCategoryCell.setValue(values);
};


function sigmaAll(){
  modifyImportedExcel();
  addCheckColumns();
  fillWithDefaultValues();
  fetchAndFillCompanyUrl();
  fillDealName();
}

function fillDealName(){

  let sheetName = source.getName().split("|");
  let dateText = sheetName[sheetName.length-1].trim();
  let eventName = sheetName[0].trim();
  let mmddyyyy = dateText.split("/");
  let eventDate = new Date(mmddyyyy[2], mmddyyyy[0]-1, mmddyyyy[1]);
  let closeDate = new Date(eventDate);
  closeDate.setDate(closeDate.getDate()-7);
  console.log(closeDate.toLocaleDateString());
  console.log(eventDate.toLocaleDateString());

  
  for(let i = 2; i <= lastRowIndex; i++){
    let companyName = destination.getRange(i, getHeaderColumnContaining("Company Name", destination)).getValue();
    let dealCell = destination.getRange(i, getHeaderColumnContaining("Deal Name", destination));
    let yearCell = destination.getRange(i, getHeaderColumnContaining("Year Of Product Or Service Pitched", destination));
    let pitchDateCell = destination.getRange(i, getHeaderColumnContaining("Pitch Call Date", destination));
    let closeDateCell = destination.getRange(i, getHeaderColumnContaining("Close Date", destination));
    dealCell.setValue(companyName + " - " + eventName + " " + dateText);
    yearCell.setValue(closeDate.getFullYear());
    pitchDateCell.setValue(closeDate.toLocaleDateString());
    closeDateCell.setValue(closeDate.toLocaleDateString());
  }
  
  
  
}


/*
* This function uses clearbit API to fetch company domain name with company and pastes it into relevant cells.
* It clearbit cannot find the company domain name, it highlights the cell red and write "Cannot find with clearbit".
*/
function fetchAndFillCompanyUrl(){
  let clearbitAutoComplete = "https://autocomplete.clearbit.com/v1/companies/suggest?query=";
  //let googleSearch = "https://www.google.com/search?q="
  let companyNames = source.getRange(2, getHeaderColumnContaining("Organization", source), numOfRecords).getValues();
  
  for(let i = 0; i < companyNames.length; i++){
    let query = clearbitAutoComplete + encodeURIComponent(companyNames[i][0]);
    let response = UrlFetchApp.fetch(query);
    let json = JSON.parse(response.getContentText());
    
    //This is where the web address will go.
    let cell = destination.getRange(i+2, getHeaderColumnContaining("Company Domain Name", destination));
    
    if(json == null || json == "") {
      cell.setValue("Cannot find with clearbit");
      cell.setBackground("red");
      cell.setFontSize(11);
      cell.setFontFamily("Calibri");
      continue;
    }

    let companyDomainName = json[0]["domain"];
    cell.setValue(companyDomainName);
    cell.setFontSize(11);
    cell.setFontFamily("Calibri");
  }
  
}



/*
* This function just fills up the destination sheet with default values for Company owner, Contact owner, 
* Attendee type, Lifecycle stage, Lead status, Pipeline, Deal Stage, Deal owner, Product or Service Pitched,
* Deal type.
*/
function fillWithDefaultValues(){


  let columnTitles = ["Company Owner","Contact Owner","Attendee Type","Lifecycle Stage","Lead Status","Pipeline","Deal Stage","Deal Owner","Product Or Service Pitched","Deal Type"];
  let columnValues = ["No owner","Tyron McGurgan","Delegate","Opportunity","Connected","Delegate Pipeline","Closed won","Tyron McGurgan","Private Events","New Business"]

  for(let i = 0; i < columnTitles.length; i++){
    let range = destination.getRange(2, getHeaderColumnContaining(columnTitles[i], destination), numOfRecords);
    range.setValue(columnValues[i]);
  }

}

/*
* This function modifies the imported excel file's header and its content appropriate to new header.
*/
function modifyImportedExcel(){
  //Select the column with the following header.
  let cols = ["Organization", "First Name", "Last Name", "Position", "Mobile Number", "Primary Email"]
  let colIndexes = [];
  for(let i = 0; i < cols.length; i++){
    colIndexes.push(getHeaderColumnContaining(cols[i], source));
  }

  //Select the range for all records. Store it to a variable.
  let ranges = []
  for(let i = 0; i < cols.length; i++){
    ranges.push(source.getRange(2, colIndexes[i], numOfRecords));
  }

  //Copy company names to fourth column where new header has "Company Name".
  let destinationColIndexes = [4, 8, 9, 11, 18, 7];
  for(let i = 0; i < ranges.length; i++){
    if(i == 4){
      ranges[i].copyValuesToRange(destination, 12, 12, 2, lastRowIndex);
    }
    ranges[i].copyValuesToRange(destination, destinationColIndexes[i], destinationColIndexes[i], 2, lastRowIndex);
  }

}


/*
* Get column number containing the string in the source Sheet.
*/
function getHeaderColumnContaining(attribute, sheet){
  let lastColumnIndex = getLastColumnIndex(sheet, 1);

  for(let i = 1; i <= lastColumnIndex; i++){
    if(sheet.getRange(1,i).getValue() == attribute){
      return i;
    }
  }
}


/*
* Find the last record in the table and return its row index.
*/
function getLastRowIndex(sheet, columnNum){

  let lastRecordRowIndex = 1;
  let foundLastRecordRowIndex = false;

  while(foundLastRecordRowIndex == false){
    if(sheet.getRange(lastRecordRowIndex,columnNum).getValue() == ""){
      foundLastRecordRowIndex = true;
      break;
    }
    lastRecordRowIndex++;
  }
  return lastRecordRowIndex-1;
}

/*
* Find the index of last column containing value in the header and return its column index.
*/
function getLastColumnIndex(sheet, rowNum){

  let lastRecordColumnIndex = 1;
  let foundLastRecordColumnIndex = false;

  while(foundLastRecordColumnIndex == false){
    if(sheet.getRange(rowNum, lastRecordColumnIndex).getValue() == ""){
      foundLastRecordColumnIndex = true;
      break;
    }
    lastRecordColumnIndex++;
  }
  return lastRecordColumnIndex-1;
}


/*
* This function adds new header containing all the fields required for Hubspot import.
*/
function addCheckColumns() {

  let newheader = ["Company Added", "Contact Added", "Deal Added", "Company Name", "Company Domain Name", "Company Owner", "Email", "First Name", "Last Name", "Contact Owner", "Job Title", "Phone Number", "Attendee Type", "Lifecycle Stage", "Lead Status", "State/Region", "Country/Region", "Mobile Phone Number", "Deal Name", "Pipeline", "Deal Stage", "Close Date", "Deal Owner", "Deal Type", "Product Or Service Pitched", "Private Event Name And Date", "Year Of Product Or Service Pitched", "Pitch Call Date", "Deal Industry", "Deal Industry Category"];
  for(let i = 0; i < newheader.length; i++){
    let currentCell = destination.getRange(1,i+1);
    currentCell.setValue(newheader[i]);
    currentCell.setBackground('#7a7a7a');
    currentCell.setFontColor('white');
  }
  
}
