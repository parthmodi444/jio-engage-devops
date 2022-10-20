const http = require('http')
const fs = require('fs')
console.log("start");
var xlsx=require("xlsx");
var dataPathExcel="final_report1.xlsx";

var wb=xlsx.readFile(dataPathExcel);
var sheetName=wb.SheetNames[0];
console.log("step 2")
var sheetValue=wb.Sheets[sheetName];
var excelData=xlsx.utils.sheet_to_json(sheetValue);
console.log(excelData);

var result = [];

for(var i in excelData)
    result.push([i, excelData[i]]);

console.log(result)

const server = http.createServer((req, res) => {
  res.writeHead(200, { 'content-type': 'text/html' })
  fs.createReadStream('index.html').pipe(res)
})

server.listen(process.env.PORT || 3000)
