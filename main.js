/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////// FUNCTIONS
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function LastUpdate(){
  
      ////// SHOW DATE WITH FORMAT IN SPREADSHEET 

  var formated_now = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");
  
  return "Last Update: " + formated_now;
}

function dateExist(_month, _year){
  
  // Check wheter a date has sense or not (if it is in the future)
  
  var int_Month = +_month;
  var int_year = +_year;
  
  var now = new Date();
  var requestedDate = new Date(int_year, int_Month - 1);
  
  if (requestedDate <= now){ // This does not let you ask about future time, just present or past
    
    return true;    
    
  } else {
    
    return false;
  }
}

function JustAccessSpreadSheet(_spreadSheetName){
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = activeSpreadsheet.getSheetByName(_spreadSheetName);
  
  return Sheet;
}

function Semrush_API(_keyword, _settingsTab, _field, _month, _year){
  
  var ThisProjectSheet = JustAccessSpreadSheet(_settingsTab);
  var PreNameCell = ThisProjectSheet.getRange("B10");
  var PreName = PreNameCell.getValue();
  var _FileName_toLookUp = PreName + _month + "-" + _year + ".json";
  
  var URLFileSheet = "URLs_" + _settingsTab;
  
  var url, cell;
  var fileAvailable = false;
  var sheet = JustAccessSpreadSheet(URLFileSheet); // This functions return the Spread Sheet with ALL File URLs
  var columns = sheet.getLastRow() + 1;
  var rows = sheet.getLastColumn() + 1;
  var selection = sheet.getRange(1,1, rows, columns).getValues();  
  
  // Loop to find my file in URL SpreadSheet
  for (var row=0; row < rows; row++) {
    
    cell = selection[row][0].toString();
    
    if (_FileName_toLookUp == cell) {
      url = selection[row][1].toString();
      fileAvailable = true;
    }
  }
  
  if (fileAvailable == true){
    
    if (dateExist(_month, _year) == true){
      
      var jsondata = UrlFetchApp.fetch(url, null);  
      var object = JSON.parse(jsondata.getContentText());
      
      var _fieldList = _field.split("/");
      var inner_array = [];
      
      if (_field == "") {
        
        // Return everything in case _field is empty
        
        inner_array.push(GetCPC(_keyword, object));
        inner_array.push(GetSearchVolume(_keyword, object));
        inner_array.push(GetRankings(_keyword, object));
        inner_array.push(GetDiffPeriod(_keyword, object));
        inner_array.push(GetDiff1day(_keyword, object));
        inner_array.push(GetDiff1week(_keyword, object));
        inner_array.push(GetDiff1month(_keyword, object));
           
      } else {
        
        for (i=0; i < _fieldList.length; i++) {
          
          switch (_fieldList[i].toLowerCase()) {
              
            case "rankings":
              inner_array.push(GetRankings(_keyword, object));
              break;
              
            case "cpc":
              inner_array.push(GetCPC(_keyword, object));
              break;
              
            case "searchvolume":
              inner_array.push(GetSearchVolume(_keyword, object));
              break;
              
            case "diffperiod":
              inner_array.push(GetDiffPeriod(_keyword, object));
              break;
              
            case "diffday":
              inner_array.push(GetDiff1day(_keyword, object));
              break;
              
            case "diffweek":
              inner_array.push(GetDiff1week(_keyword, object));
              break;
              
            case "diffmonth":
              inner_array.push(GetDiff1month(_keyword, object));
              break;
              
            default:
              inner_array.push("Please use a correct Field (rankings, cpc, searchvolume, diffperiod, diffday, diffweek, diffmonth). You can leave it empty, between quotes, to get all parameters");
              
          }
        }
      }
        // This transformation returns a row instead of a column, like doing a matrix transpose in the spreadsheet
      var output_array = [];
      output_array.push(inner_array);
      
      return output_array;
      
    } else {
      
      return "Sorry, there is no data available for this date";
    }
  } else { return "File not available, Download it manually and then Refresh URLs SpreadSheets"; }
}

function GetRankings(_keyword, object){     
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){
      
      for (var key in object.data[i].Fi) {
        var output = object.data[i].Fi[key];
        if(output == "-"){output = 101;}
        
        //return object.data[i].Fi[key]; // Position at the end of specified period
        return output;
      }
    }
  }
}

function GetCPC(_keyword, object){
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){ 
      var output = object.data[i].Cp;
      if(output == "n/a"){output = 0;}
      
      //return object.data[i].Cp; // average price (US dollars)
      return output; 
    }
  }  
}

function GetSearchVolume(_keyword, object){
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){
      
      var output = object.data[i].Nq ;
      if(output == "n/a"){output = 0;}
      
      //return object.data[i].Nq; // average search per month by user (last 12 months)        
      return output; 
    }
  }
}

function GetDiffPeriod(_keyword, object){
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){
      
      for (var key in object.data[i].Fi) {
        
        var output = object.data[i].Diff[key];
        if(output == "-"){output = 0;}
        
        //return object.data[i].Diff[key]; // diff (specified period)
        return output; 
      }
    }
  }
}

function GetDiff1day(_keyword, object){
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){
      
      for (var key in object.data[i].Fi) {
        
        var output = object.data[i].Diff1[key];
        if(output == "-"){output = 0;}
        
        return output; 
        //return object.data[i].Diff1[key]; // diff by day
      }
    }
  }
}

function GetDiff1week(_keyword, object){
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){
      
      for (var key in object.data[i].Fi) {
        
        var output = object.data[i].Diff7[key];
        if(output == "-"){output = 0;}
        
        return output; 
        //return object.data[i].Diff7[key]; // diff by week
      }
    }
  }
}

function GetDiff1month(_keyword, object){
  
  // Keywords lenght
  var keywords_len = Object.keys(object.data).length;
  
  // Match every row with corresponding keyword
  
  for(var i=0; i<keywords_len; i++ ) {
    
    if (_keyword == object.data[i].Ph){
      
      for (var key in object.data[i].Fi) {
        
        var output = object.data[i].Diff30[key];
        if(output == "-"){output = 0;}
        
        return output; 
        //return object.data[i].Diff30[key]; // diff by month
      }
    }
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////
//////////////////    BUILD SPREADSHEET CONTAINING ALL JSON URLS    
///////////////////////////////////////////////////////////////////////////////////////////////

function BuildURLSpreadSheet(SettingsTabName){
    
  var URLFilesTabName = "URLs_" + SettingsTabName;
   
  // Write All URLs into a matrix (data_array)
  
  var ProjectSheet = JustAccessSpreadSheet(SettingsTabName);
  var PreName_Cell = ProjectSheet.getRange("B10");
  var PreName_value = PreName_Cell.getValue();
  var ProjectFolder_Cell = ProjectSheet.getRange("G11");
  var ProjectFolder = ProjectFolder_Cell.getValue();
  
  var data_array = [];
  var internal_array;
 
  var files = DriveApp.getFolderById(ProjectFolder).getFiles();
  
    while (files.hasNext()) {
      var file = files.next();
      var ThisfileName = file.getName();
      
      // Take files that contains "Pre-Name" from Settings tab
      if (ThisfileName.indexOf(PreName_value)>-1 == true){
       
        // Get JSON File URL
        var ID = file.getId();
          
        var DownloadUrl = "https://drive.google.com/uc?export=download&id=";
        var JsonFileUrl = DownloadUrl + ID;
       
        var internal_array = [];
        internal_array.push(ThisfileName); // File Name
        internal_array.push(JsonFileUrl);  // FIle URL
        
        // Add row
        data_array.push(internal_array); 
      }
    }
  
  // Get URLSheet
  var URLSheet = JustAccessSpreadSheet(URLFilesTabName);
  
  // In case the Sheet does not exist, create it
    if (URLSheet == null) {
        var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        URLSheet = activeSpreadsheet.insertSheet();
        URLSheet.setName(URLFilesTabName);
        }
  
  // Print matrix in spreadsheet (URLSheet)
  var num_rows = data_array.length;
  var num_col = data_array[0].length;
  var URLSheet_Cell = URLSheet.getRange(1, 1, num_rows, num_col);
  
  // Clear previous values
  var allRows = URLSheet.getLastRow() + 1;
  var allColumns = URLSheet.getLastColumn() + 1;
  var allSheet = URLSheet.getRange(1, 1, allRows, allColumns);
  allSheet.clearContent();
  
  // Print new values
  var URLSheet_value = URLSheet_Cell.setValues(data_array);       
       
}

///////////////////////////////////////////////////////////////////////////////////////////////
//////////////////    AUTOMATED DOWNLOAD - GET JSON FROM SEMRUSH API AND STORE IN GOOGLE DRIVE    
///////////////////////////////////////////////////////////////////////////////////////////////

function SaveJSONfromURL(_settingsTab) {
    
    ////// SAVE JSON ONCE A MONTH (from URL)  

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var SettingSheet = activeSpreadsheet.getSheetByName(_settingsTab);
  
  var FileCell = SettingSheet.getRange("D11");
  var FileNameRaw = FileCell.getValue();
  var FileName = FileNameRaw + ".json";
  
  var SEMRushAPIKey_Cell = SettingSheet.getRange("B4");
  var SEMRushAPIKey = SEMRushAPIKey_Cell.getValue();
  
  var SEMRushProjectID_Cell = SettingSheet.getRange("B7");
  var SEMRushProjectID = SEMRushProjectID_Cell.getValue();
  
  var ProjectDomain_Cell = SettingSheet.getRange("D4");
  var ProjectDomain = ProjectDomain_Cell.getValue();
  
  var StartPeriod_Cell = SettingSheet.getRange("D7");
  var StartPeriod = StartPeriod_Cell.getValue();
  
  var EndPeriod_Cell = SettingSheet.getRange("E7");
  var EndPeriod = EndPeriod_Cell.getValue();
  
  var KeywordLimit_Cell = SettingSheet.getRange("G7");
  var KeywordLimit = KeywordLimit_Cell.getValue();
  
  var ProjectFolder_Cell = SettingSheet.getRange("G11");
  var ProjectFolder = ProjectFolder_Cell.getValue();
    
     // MAIN URL
  
  var url = "http://api.semrush.com/reports/v1/projects/" + SEMRushProjectID + "/tracking/?key=" + SEMRushAPIKey + "&action=report&type=tracking_position_organic&display_limit=" + KeywordLimit + "&display_sort=cp_desc&date_begin=" + StartPeriod + "&date_end=" + EndPeriod + "&url=*." + ProjectDomain + "%2F*&display_offset=0";
  //Logger.log(url);
  var response = UrlFetchApp.fetch(url);
  var json_raw = response.getContentText();
  
     // Delete previous JSON of this month, if existed
  
  function DeleteLastJSON(JsonToDeleteName, ID) {
    var files = DriveApp.getFolderById(ID).getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (JsonToDeleteName == file.getName()){
        file.setTrashed(true);
      }
    }
  }
  
  /// PONER ESTO DENTRO DE LOS SUBFOLDERS
  DeleteLastJSON(FileName, ProjectFolder); // ProjectFolder -> ID
  
  //Save JSON in ID Folder
  DriveApp.getFolderById(ProjectFolder).createFile(FileName, json_raw, MimeType.PLAIN_TEXT);
  
}

///////////////////////////////////////////////////////////////////////////////////////////////
//////////////////    MANUAL DOWNLOAD - GET JSON FROM SEMRUSH API AND STORE IN GOOGLE DRIVE    
///////////////////////////////////////////////////////////////////////////////////////////////

function Manual_SaveJSONfromURL(_settingsTab) {
    
    ////// SAVE JSON ONCE A MONTH (from URL) 

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var SettingSheet = activeSpreadsheet.getSheetByName(_settingsTab);
  
  var FileCell = SettingSheet.getRange("B21");
  var FileNameRaw = FileCell.getValue();
  var FileName = FileNameRaw + ".json";
  
  var SEMRushAPIKey_Cell = SettingSheet.getRange("B4");
  var SEMRushAPIKey = SEMRushAPIKey_Cell.getValue();
  
  var SEMRushProjectID_Cell = SettingSheet.getRange("B7");
  var SEMRushProjectID = SEMRushProjectID_Cell.getValue();
  
  var ProjectDomain_Cell = SettingSheet.getRange("D4");
  var ProjectDomain = ProjectDomain_Cell.getValue();
  
  var StartPeriod_Cell = SettingSheet.getRange("C18");
  var StartPeriod = StartPeriod_Cell.getValue();
  
  var EndPeriod_Cell = SettingSheet.getRange("D18");
  var EndPeriod = EndPeriod_Cell.getValue();
  
  var KeywordLimit_Cell = SettingSheet.getRange("G7");
  var KeywordLimit = KeywordLimit_Cell.getValue();
  
  var ProjectFolder_Cell = SettingSheet.getRange("G11");
  var ProjectFolder = ProjectFolder_Cell.getValue();
    
     // MAIN URL
  
  var url = "http://api.semrush.com/reports/v1/projects/" + SEMRushProjectID + "/tracking/?key=" + SEMRushAPIKey + "&action=report&type=tracking_position_organic&display_limit=" + KeywordLimit + "&display_sort=cp_desc&date_begin=" + StartPeriod + "&date_end=" + EndPeriod + "&url=*." + ProjectDomain + "%2F*&display_offset=0";
  //Logger.log(url);
  var response = UrlFetchApp.fetch(url);
  var json_raw = response.getContentText();
  
     // Delete previous JSON of this month, if existed
  
  function DeleteLastJSON(JsonToDeleteName, ID) {
    var files = DriveApp.getFolderById(ID).getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (JsonToDeleteName == file.getName()){
        file.setTrashed(true);
      }
    }
  }
  
  /// PONER ESTO DENTRO DE LOS SUBFOLDERS
  DeleteLastJSON(FileName, ProjectFolder); // ProjectFolder -> ID
  
  //Save JSON in ID Folder
  DriveApp.getFolderById(ProjectFolder).createFile(FileName, json_raw, MimeType.PLAIN_TEXT);
  
}

///////////////////////////////////////////////////////////////////////////////////////////////

//                                         END

///////////////////////////////////////////////////////////////////////////////////////////////
