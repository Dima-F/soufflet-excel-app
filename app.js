const Excel = require('exceljs');
const workbook = new Excel.Workbook();
const fs = require('fs-extra');
const path = require('path');

(async () => {
    await fs.emptyDir(path.join(process.cwd(), 'output'));

    const bBook = await workbook.xlsx.readFile('book.xlsx');
    const bWorksheet = bBook.getWorksheet(1);
    const cBook = await workbook.xlsx.readFile('card.xlsx');
    const cWorksheet1 = cBook.getWorksheet('ik');
    const cWorksheet2 = cBook.getWorksheet('a');
    const outerIterator = [];

    for(let j=2;j<=413;j++) {
        outerIterator.push(j);
    }

    for(const i of outerIterator) {
        const nameCell = bWorksheet.getCell(`C${i}`).value;
        const osnZasCell = bWorksheet.getCell(`A${i}`).value + '-0';
        const invCell = bWorksheet.getCell(`D${i}`).value;
        const dateCell = bWorksheet.getCell(`B${i}`).value;
        const priceCell = bWorksheet.getCell(`E${i}`).value;
        const datadoc = bWorksheet.getCell(`N${i}`).value;
        const docnumber = bWorksheet.getCell(`M${i}`).value;
        const account = bWorksheet.getCell(`J${i}`).value;
        const codeanaccount = bWorksheet.getCell(`I${i}`).value;
        const coderesponperson = bWorksheet.getCell(`F${i}`).value;
        const posadarespperson = bWorksheet.getCell(`H${i}`).value;
        const PIB = bWorksheet.getCell(`G${i}`).value;
        const company = bWorksheet.getCell(`L${i}`).value;

        
        cWorksheet1.getCell('B9').value = nameCell;
        cWorksheet1.getCell('C20').value = nameCell;
        cWorksheet1.getCell('N20').value = osnZasCell;
        cWorksheet1.getCell('B26').value = invCell;
        cWorksheet1.getCell('N54').value = dateCell;
        cWorksheet1.getCell('M20').value = dateCell;
        cWorksheet1.getCell('N13').value = priceCell;
        cWorksheet1.getCell('K13').value = datadoc;
        cWorksheet1.getCell('M13').value = docnumber;
        cWorksheet1.getCell('B20').value = account; 
        cWorksheet1.getCell('E20').value = codeanaccount; 
                

        cWorksheet2.getCell('M12').value = docnumber;
        cWorksheet2.getCell('O12').value = datadoc;
        cWorksheet2.getCell('Q12').value = coderesponperson;
        cWorksheet2.getCell('C20').value = account;
        cWorksheet2.getCell('G20').value = priceCell;
        cWorksheet2.getCell('H20').value = invCell;
        cWorksheet2.getCell('J20').value = codeanaccount;
        cWorksheet2.getCell('S20').value = dateCell;        
        cWorksheet2.getCell('C24').value = nameCell;
        cWorksheet2.getCell('E25').value = company;
        cWorksheet2.getCell('C29').value = nameCell;
        cWorksheet2.getCell('B67').value = posadarespperson;
        cWorksheet2.getCell('G67').value = PIB;     
        cWorksheet2.getCell('F72').value = dateCell;
       
        await cBook.xlsx.writeFile(`./output/${osnZasCell}.xlsx`);
    }
})().then(()=>console.log('Process finished...'));


