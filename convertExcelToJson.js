var xlsx=require("xlsx")
var fs=require('fs')
var dataPathExcel="Market data.xlsx"
var wb=xlsx.readFile(dataPathExcel)
for(let i=0;i<wb.SheetNames.length;i++){
    var sheetName=wb.SheetNames[i];
    var sheetvalue=wb.Sheets[sheetName];
    var excelData=xlsx.utils.sheet_to_json(sheetvalue)
    console.log(excelData)
    console.log("---------")
    fs.writeFile(sheetName+".json",JSON.stringify(excelData),function(err){
    
        console.log("Json file Created")
    })
}