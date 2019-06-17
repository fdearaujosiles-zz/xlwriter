const WriteXLSX = require('../lib/Workbook')

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

    console.log(sheet.cell('A2').value, '\n', sheet.cell('A2'))
    

    sheet.cell('A2').value = '    ="Novo texto"'
    
    console.log(sheet.cell('A2').value, '\n', sheet.cell('A2'))
    
    console.log('\n')
    
    xls.sheet('Sheet1').range('A1:B4').conditionalFormatting('"Name"', formattingStyle);
    xls.sheet('Sheet1').range('A1:B4').conditionalFormatting('"Fernando"', formattingStyle);

    xls.save('OutTest1.xlsx', () => console.log('saved'))
}).catch(err => console.log(err));