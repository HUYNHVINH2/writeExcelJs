const Excel = require('exceljs');
let filename = 'input.xlsx';

let workbook = new Excel.Workbook();
var indexR ;
var dataAppend = [
    {name:'nguyen van A', old :'15' },
    {name:'nguyen van B', old :'20' },
    {name:'nguyen van C', old :'22' }
];
async function appendDataExecl(name, old){
    try {
        let data = await workbook.xlsx.readFile(filename);
        let indexR = await data.worksheets[0].rowCount ;
        console.log(indexR)
        await data.worksheets[0].addRow([name, old]);//get firt ws,and add new row 
        await workbook.xlsx.writeFile(filename);
        converCss(indexR+1)
    } catch (error) {
        console.log(error);
    }
}
async  function exListRows (){
    for(let i = 0 ; i<dataAppend.length ;i++){
    
    await appendDataExecl(dataAppend[i].name, dataAppend[i].old);
   
}
}

async function converCss(currenIndex){
    try {
        console.log( currenIndex)
        let data = await workbook.xlsx.readFile(filename);
        data.worksheets[0].getRow(Number(currenIndex)).font = {
                name: 'Calibri',
                color: { argb: 'FFFF0000' },
                family: 2,
                size: 14,
                italic: true
            };
            await workbook.xlsx.writeFile(filename);
    } catch (error) {
        console.log(error);
    }
}
exListRows();
