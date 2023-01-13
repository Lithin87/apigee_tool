const {XMLParser} = require('fast-xml-parser');
var jp = require('jsonpath');
var fs = require('fs');
var xl = require('excel4node');

    const options = {
        ignoreAttributes: true,
        attributeNamePrefix : "@",
        removeNSPrefix: true
    };
    
var xmlDataStr = fs.readFileSync("Source.txt");

const parser = new XMLParser(options);
const output = parser.parse(xmlDataStr);
    
var a1=[]
var a2=[]
var a3=[]
var a4=[]
     
a1.push(...jp.query(output, '$..apiSessionId'));
a2.push(...jp.query(output, '$..qflowCustomerId'));
a3.push(...jp.query(output, '$..remoteId'));
a4.push(...jp.query(output, '$..remoteType'));
    

fs.rmSync("Final.txt");
fs.rmSync("Apigee.xlsx");


  for (var i in a1) 
  {
  var t= "\n"+a1[i]+ ","+ a2[i] +","+ a3[i] + ","+ a4[i];
  fs.writeFileSync("Final.txt",t,{flag:'a'})

}
const df = fs.readFileSync("Final.txt");
const lines = df.toString().split(/\r?\n/);


var wb = new xl.Workbook();
var ws = wb.addWorksheet('TAB 1');
for (let j=1;j<=lines.length-1;j++){             
    // Create a reusable style
    var style = wb.createStyle({
      font: {
        color: '#050000',
        size: 12,
      },
  
    });
  
      pieces = lines[j].split(",")
      pieces.forEach((element, index) =>{
      ws.cell(j, index+1).string(element).style(style);
      });
  
  } 
wb.write("Apigee.xlsx");;
