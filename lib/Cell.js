const Range = require('./Range');

const {parse} = require('./elementtree');

class Cell {
    constructor(xml, sheet) {
        this.et = parse(xml)
        this.Sheet = sheet
    }


}

module.exports = Cell