const {SubElement, appendElements} = require('./elementtree');

class Cell {
    constructor(xml=undefined, sheet=undefined) {
        if(sheet === undefined) throw Error('New Cell must be appended to a Sheet.');
        if(typeof xml === 'string') {
            if(!xml.match(/^[A-z]{1,}(\d){1,}$/)) throw Error('New Cell must have a elementtree or reference.');
            this.ref = xml;
            this._value = '';
            let sheetData = sheet.et.find('sheetData')
            let row = sheetData.find('row[@r="${xml}"]') || SubElement(sheetData, 'row', {r: xml})
            appendElements(row, {
                c: {
                    attrib: {r: xml},
                    v: {text: ''}
                }
            })
            this.et = row.find(`c[@r="${xml}"]`);
        } else {
            this.et = xml;
            this.ref = this.et.attrib.r;
            if(this.et.attrib.t == "s") this._value = sheet.Workbook.sharedStrings.findall('si/t')[this.et.find('v').text].text
            else {
                if (xml.find('f')) this._value = '=' + xml.find('f').text
                else this._value = (xml.find('v') || {text: ''}).text
            }
        }
        this.Sheet = sheet;
    }
    
    get value() {return this._value}

    set value(newValue) {
        newValue = newValue.trim();
        this._value = newValue;
        if(typeof newValue != 'string' || newValue[0] == '=') {
            delete this.et.attrib.t;
            if(newValue[0] == '=') {
                if(this.et.find('f')) this.et.find('f').text = newValue.replace(/^=/, '');
                else appendElements(this.et, {f: {text: newValue.replace(/^=/, '')}});
                if(this.et.find('v')) this.et.remove(this.et.find('v'))
            }
            else this.et.find('v') ? this.et.find('v').text = newValue : appendElements(this.et, {v: {text: newValue}})
        } else {
            this.et.attrib.t = 's';
            let sst = this.Sheet.Workbook.sharedStrings.findall('si/t');
            let i = sst.filter(f => f.text == newValue);
            if(i.length) this.et.find('v') ? this.et.find('v').text = sst.indexOf(i[0]) : appendElements(this.et, {v: {text: sst.indexOf(i[0])}})
            else {
                this.Sheet.Workbook.addSST(newValue);
                let sstId = this.Sheet.Workbook.sharedStrings.getroot().attrib.uniqueCount - 1;
                this.et.find('v') ? this.et.find('v').text = sstId : appendElements(this.et, {v: {text: sstId}})
            }
        }
    }
}

module.exports = Cell