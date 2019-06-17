const Cell = require('./Cell');
const Range = require('./Range');

const {parse} = require('./elementtree');

class Sheet {

    constructor(xml, wb) {
        this.et = parse(xml);
        this.Workbook = wb;
        this.et.findall('*/row/c').forEach(cell => this.cells)
    }

    // get cells() {
    //     let x = this.et.findall('*/row/c')
    //     return x.reduce((a,b) => a.concat(b), [])
    // }

    // cell(ref) {

    // }

    // range(ref) {

    // }

    conditionalFormatting(ref, formula, style, type='cellIs', operator='equal') {
        let dxfs = this.Workbook.style.find('dxfs');

        this.et.appendElements(dxfs, {dxf: style});
        this.et.appendElements(this.et.find('./'), {
            'conditionalFormatting': {
                attrib: {sqref: ref},
                cfRule: {
                    attrib: {
                        type: type,
                        dxfId: String(parseInt(dxfs.attrib.count)-1),
                        priority: "1",
                        operator: operator
                    },
                    formula: {text: formula}   
                }
            }
        })
        return this.organize()
    }

    organize() {
        this.et.find('./')._children.sort((a,b) => {
            let orderBy = ['dimension',
            'sheetViews',
            'sheetFormatPr',
            'sheetData',
            'conditionalFormatting',
            'pageMargins',
            'pageSetup',
            'ignoredErrors']
            return orderBy.indexOf(a.tag) > orderBy.indexOf(b.tag)
        })
        return this
    }

    write(obj) {return this.et.write(obj)}
}

module.exports = Sheet