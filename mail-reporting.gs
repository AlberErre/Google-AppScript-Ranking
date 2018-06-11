function RankingNewsControl() {
  
  // CONFIG Email variables
  
    // Campaigns Sheet - include all project sheets manually
    //var RankingSheets_list = ["Rankings_company_1, Rankings_company_2, Rankings_company_3"]
  
  var email_cols = 4, email_init_col = 1;
  var email_addresses, email_subject, email_content;
  var email_send_spent, email_send_ratio;
  
  // Spreadsheets variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var keyword, ranking, diffperiod, lastPeriod_ranking, diffperiod_available;
  var spreadSheet_name, spreadSheet_name_and_keyword, rankingColumn, diffColumn;
  var email_send_enterTop100, email_send_enterTop10, email_send_movementTop10;
  
  //Email matrix data
  var top_100_keywordMatrix, top_10_keywordMatrix, top_10Movements_keywordMatrix;
  
  // RankingSheets_list  
  var x_sheet = ss.getSheetByName("Mail-Reporting-Settings");
  var x_rows = x_sheet.getLastRow();  
  var RankingSheets_array = x_sheet.getRange(2, 1, x_rows, 1).getValues();
  var RankingSheets_list = [];
  for (var item=0; item<(RankingSheets_array.length -1); item++){
    RankingSheets_list.push(RankingSheets_array[item][0]);    
  }
  
  Logger.log(RankingSheets_list);
    
  // Control Sheet
  var c_sheet = ss.getSheetByName("Mail-Reporting-Settings");
  var c_cols = c_sheet.getLastColumn();
  var c_rows = c_sheet.getLastRow();
  var c_data = c_sheet.getRange(1, 1, c_rows, c_cols).getValues();
  var days;
  
  // Project Sheets  
  for (var i=0; i<RankingSheets_list.length; i++){
    var p_sheet = ss.getSheetByName(RankingSheets_list[i]);
    Logger.log(p_sheet);
    var cols = p_sheet.getLastColumn();
    var rows = p_sheet.getLastRow();  
    var data = p_sheet.getRange(1, 1, rows, cols).getValues();
    
    //Logger.clear();
    
    for (var c_row=1; c_row<c_rows; c_row++) {
      
      // Email Addresses
      email_addresses = [];
      
      // Generate Email List
      for (var row2=0; row2<email_cols; row2++) {
        
        if (c_data[c_row][email_init_col+row2] != "") {
          email_addresses.push(c_data[c_row][email_init_col+row2]);
        } else { break; }
      }
      
      // Lookup column index for ranking and diff period
      for (var x=0; x<cols;x++){
        if (data[1][x] == "Ranking"){rankingColumn = x}
        if (data[1][x] == "Diff Period"){diffColumn = x}
      }      
          
      top_100_keywordMatrix = [];
      top_10_keywordMatrix = [];
      top_10Movements_keywordMatrix = [];
      
      for (var row=2; row<rows; row++) {
        
        // Get Campaign Data
        keyword = data[row][0];
        ranking = data[row][rankingColumn];
        diffperiod = data[row][diffColumn];
        
        lastPeriod_ranking = ranking + diffperiod;
        
        email_send_enterTop100 = false;
        email_send_enterTop10 = false;
        email_send_movementTop10 = false;
        
        spreadSheet_name = RankingSheets_list[i];
        spreadSheet_name_and_keyword = spreadSheet_name + "" + keyword;
        
        if (keyword !== "" && ranking !== "" && ranking !== "-" ) {          
          
          // Conditions  
          // Enter top 100
          if (lastPeriod_ranking > 100 && ranking <= 100){
            
            email_send_enterTop100 = true;
            
            // PUSH ROW WITH DATA IN MATRIX
            var this_keyword100 = "Spreadsheet:  " + "<strong>" + spreadSheet_name + "</strong>" + ",  Keyword:  " + "<strong>" + keyword + "</strong>" + ".";
            
            var internalArray100 = [];
            internalArray100.push("<p style='margin-left: 25px;'>" + this_keyword100 + "</p>");
            top_100_keywordMatrix.push(internalArray100);
          }
          
          // Enter top 10
          else if (lastPeriod_ranking > 10 && ranking <= 10){
            
            email_send_enterTop10 = true;
            
            // PUSH ROW WITH DATA IN MATRIX
            var this_keyword10 = "Spreadsheet:  " + "<strong>" + spreadSheet_name + "</strong>" + ",  Keyword:  " + "<strong>" + keyword + "</strong>" + ".";
            
            var internalArray10 = [];
            internalArray10.push("<p style='margin-left: 25px;'>" + this_keyword10 + "</p>");
            top_10_keywordMatrix.push(internalArray10);
            
          }
          
          // Movements inside top 10
          else if (lastPeriod_ranking <= 10 && ranking <= 10 && ranking != lastPeriod_ranking){
            
            email_send_movementTop10 = true;
            
            // PUSH ROW WITH DATA IN MATRIX
            var this_keyword10Mov = "Spreadsheet:  " + "<strong>" + spreadSheet_name + "</strong>" + ",  Keyword:  " + "<strong>" + keyword + "</strong>" + 
                                    ". From position:  " + lastPeriod_ranking + "  to position:  " + ranking ;
            
            var internalArray10mov = [];
            internalArray10mov.push("<p style='margin-left: 25px;'>" + this_keyword10Mov + "</p>");
            top_10Movements_keywordMatrix.push(internalArray10mov);
            
          }
        }           
      }

      //Send Emails
      
      if (email_addresses.length > 0 && email_send_enterTop100 == true) {
        
        // TOP 100
        email_subject = "[Semrush Ranking info] - Some keywords have reached top 100";
        email_content = "<p>Hello,</p>" + 
          "<p>This is a Semrush Auto-Rankings tool report, the following keywords have reached top 100: " +
            "<p style='margin-left: 25px;'><strong>List of keywords:</strong></p>" +
               top_100_keywordMatrix +
                "<p></p><p></p>" +
                  "<p><i>Semrush Auto-Rankings tool</i></p>";
        
        SendEmail(email_addresses, email_subject, email_content);
      } 
      
      if (email_addresses.length > 0 && email_send_enterTop10 == true) {
        
        // TOP 10
        email_subject = "[Semrush Ranking info] - Some keywords have reached top 10";
        email_content = "<p>Hello,</p>" + 
          "<p>This is a Semrush Auto-Rankings tool report, the following keywords have reached top 10: " +
            "<p style='margin-left: 25px;'><strong>List of keywords:</strong></p>" +
               top_10_keywordMatrix +
                "<p></p><p></p>" +
                  "<p><i>Semrush Auto-Rankings tool</i></p>";
        
        SendEmail(email_addresses, email_subject, email_content);
      } 
      
      if (email_addresses.length > 0 && email_send_movementTop10 == true) {
        
        // TOP 10 MOVEMENTS
        email_subject = "[Semrush Ranking info] - Some movements have took place within top 10 keywords";
        email_content = "<p>Hello,</p>" + 
          "<p>This is a Semrush Auto-Rankings tool report, the following keywords have changed their positions within the top 10: " +
            "<p style='margin-left: 25px;'><strong>List of keywords:</strong></p>" +
              top_10Movements_keywordMatrix +
                "<p></p><p></p>" +
                  "<p><i>
        's Semrush Auto-Rankings tool</i></p>";
        
        SendEmail(email_addresses, email_subject, email_content);
      }
      
     } //endfor
                  
  } //endfor list
  
} //endfunction


function SendEmail(email_to, email_subject, email_content) {
  
  MailApp.sendEmail({
    name: "Semrush auto-ranking",
    to: email_to.join(","),              
    subject: email_subject,
    htmlBody: email_content 
  });
  
}
