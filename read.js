var xlsx=require("xlsx");
var dataPathExcel="jio.xlsx";
var wb=xlsx.readFile(dataPathExcel);
var sheetName=wb.SheetNames[2];
var sheetValue=wb.Sheets[sheetName];
var excelData=xlsx.utils.sheet_to_json(sheetValue);
console.log(excelData);
