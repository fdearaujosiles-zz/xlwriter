const Cell = require('./Cell');
const Range = require('./Range');

const {parse, appendElements} = require('./elementtree');

class Sheet {

    constructor(xml, wb) {
        this.Workbook = wb;
        this.et = parse(xml);
        this.cells = this.et.findall('*/row/c').map(cell => new Cell(cell, this));
    }

    cell(ref) {return this.cells.filter(c => c.ref.toLowerCase() == ref.toLowerCase())[0] || new Cell(ref, this)}

    range(ref) {return new Range(ref, this)}

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