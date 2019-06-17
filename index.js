const WriteXLSX = require('./lib/Workbook')

WriteXLSX.read("../out.xlsx").then(xls => {

    // console.log(fs.readFileSync('assets/out/xl/workbook.xml').toString())
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

    console.log(sheet.cells)
    
    // xls.sheet('Sheet1').conditionalFormatting('A1:B4', '"Name"', formattingStyle);
    // xls.sheet('Sheet1').conditionalFormatting('A1:B4', '"Fernando"', formattingStyle);

    xls.save('OutTest1.xlsx', () => console.log('saved'))
}).catch(err => console.log(err));