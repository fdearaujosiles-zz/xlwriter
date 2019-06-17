const WriteXLSX = require('../lib/Workbook')


let xls = new WriteXLSX();
console.log('\n', xls.sharedStrings)
xls.sheet('Sheet1').cell('A1').value = 'NewText';
console.log('\n', xls.sharedStrings)
console.log('\n', xls.sheet(0).cell('A1').et)


WriteXLSX.read("../out.xlsx").then(xls => {

    formattingStyle = {
        font: {
            b:true, 
            i: true,
            stroke: true, 
            color: {
                attrib: {
                    rgb: 'FFDDDDDD'
                }
            }
        },
        fill: {
            patternFill: {
                bgColor: {
                    attrib: {
                        rgb:'FFDDDDDD'
                    }
                }
            }
        }
    };

    let sheet = xls.sheet('Sheet1');
    let joy = xls.sheet('@joy');

    console.log('\n')

    console.log(sheet.cells.map(cell => cell.value))
    sheet.cell('a3').value = 64
    console.log(sheet.cells.map(cell => cell.value))
    
    console.log('\n')
    
    // xls.sheet('Sheet1').range('A1:B4').conditionalFormatting('"Name"', formattingStyle);
    
    xls.save('OutTest1.xlsx', () => console.log('saved'))
}).catch(err => console.log(err));