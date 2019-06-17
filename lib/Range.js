const {appendElements} = require('elementtree') 

class Range {
    constructor(ref, sheet) {
        if(sheet === undefined) throw Error('New Cell must be appended to a Sheet.');
        this.Sheet = sheet;
        ref = ref.trim()
        if(!ref.match(/^([A-z]{1,}(\d){1,}:[A-z]{1,}(\d){1,}||[A-z]{1,}:[A-z]{1,}||(\d){1,}:(\d){1,})$/))
            throw Error('Range reference string must be in the format "A1:B1"');
        this.ref = ref;
    }

    conditionalFormatting(formula, style, type='cellIs', operator='equal') {
        let dxfs = this.Sheet.Workbook.style.find('dxfs');

        appendElements(dxfs, {dxf: style});
        appendElements(this.Sheet.et.find('./'), {
            'conditionalFormatting': {
                attrib: {sqref: this.ref},
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
        return this.Sheet.organize()
    }

}

module.exports = Range