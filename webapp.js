function doGet() {
  var index_template = HtmlService.createTemplateFromFile("index");
  
  index_template.output_text = "";
 
  var index_evaluation = index_template.evaluate();
  index_evaluation.setTitle('レンタサイクル監視');
//  console.log(index_evaluation.getContent());
  return index_evaluation;
}

function doPost(e){
  console.log(e);
//  var input_text = e.parameter.input;
//  console.log(input_text);
  var index_template = HtmlService.createTemplateFromFile("index");
  
  var output_text = "押されました";
  
  index_template.output_text = output_text;
  return index_template.evaluate(); 

//  return ContentService.createTextOutput('入力した値は'+ input_text);
  
}

function createTable(sheet_name){
  var spread_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  let count = spread_sheet.getLastRow();
  let column_count = 3;
  var values;
  if(count > 0){
    values = spread_sheet.getRange(1, 1, count, column_count).getValues();
  }else{
    values = [Array(column_count).fill("")];
  }
  
  var table_html = "";
  table_html = "<table><tr><th>Type</th><th>bike_no</th><th>Date</th></tr>";
  for(var row in values){
    table_html += "<tr>";
    for(var column in values[row]){
      table_html += "<td>" + values[row][column] +"</td>";
    }
    table_html += "</tr>";    
  }
  
  table_html += "</table>";
  
  return table_html;
   
}

