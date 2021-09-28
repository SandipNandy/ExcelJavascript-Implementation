var XLSX = require('xlsx');
var workbook = XLSX.readFile('C:/Users/SandipNandi/Project/environment_details.xlsx');
var worksheet=workbook.Sheets['Sheet2'];
var data=XLSX.utils.sheet_to_json(worksheet);

var NewJSONData=data;

  for (let len = Object.keys(NewJSONData).length, i = len; i < len * 3; i++) {
    Object.assign(NewJSONData, {[i]: JSON.parse(JSON.stringify(NewJSONData[i -1]))});
  }
  
var newworkbook=XLSX.utils.book_new();
var newWS= XLSX.utils.json_to_sheet(NewJSONData);
XLSX.utils.book_append_sheet(newworkbook,newWS,'New_DATA');
XLSX.writeFile(newworkbook,"NewDataFile.xlsx");
const fs = require('fs');
for(let i=1;i<=10;i++){
fs.copyFile('C:/Users/SandipNandi/Project/NewDataFile.xlsx', 'C:/Users/SandipNandi/Desktop/XL_COPY/NewDataFileCopied_'+i+'.xlsx', (err) => {
  if (err) throw err;
});
}
