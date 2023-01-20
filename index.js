const {XMLParser} = require('fast-xml-parser');
var jp = require('jsonpath');
var fs = require('fs');
var xl = require('excel4node');

    const options = {
        ignoreAttributes: true,
        removeNSPrefix: true
    };
    
const parser = new XMLParser(options);
const output = parser.parse(fs.readFileSync("Source.txt"));
    

var a1 = jp.query(output, '$..apiSessionId');
var a2 = jp.query(output, '$..qflowCustomerId');
var a3 = jp.query(output, '$..remoteId');
var a4 = jp.query(output, '$..time');
var a5 = jp.query(output, '$..status');

var file_output = jp.query(output , '$..*')

fs.writeFileSync("XMLParser",JSON.stringify(file_output),{flag:'a'})

fs.rmSync("Final.txt", { force: true });
fs.rmSync("Apigee.xlsx", { force: true });

var t= "Session ID"+ ","+ "QflowID" +","+ "CustomerID" + ","+ "Time" + "," + "Pass/Fail";
for (var i in a1) 
{ 
  fs.writeFileSync("Final.txt",t,{flag:'a'})
   var t= "\n"+a1[i]+ ","+ a2[i] +","+ a3[i] + ","+ a4[i] + "," + a5[i];
}

var wb = new xl.Workbook();
var ws = wb.addWorksheet('TAB 1');

// var style = wb.createStyle({
//   font: {
//     color: '#050000',
//     size: 14,
//   },
// });

const lines = fs.readFileSync("Final.txt").toString().split(/\r?\n/);
lines.forEach((c,i) =>  c.split(",").forEach((element, index) =>  ws.cell(i+1, index+1).string(element)) )

wb.write("Apigee.xlsx");
