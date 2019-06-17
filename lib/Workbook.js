const Sheet = require('./Sheet');

const JSZip = require("jszip");
const walkSync = require('./walkSync');
const {parse, appendElements} = require('elementtree');
const {readFileSync, createWriteStream} = require('fs');

class Workbook {
    constructor(template=undefined) {
        if(template === undefined) {
            template = __dirname + '/../assets/template/';
            walkSync(template).forEach(file => {
                if(file.match(/sheet(\d){1,}.xml$/)) this[file.replace(template, '')] = new Sheet(readFileSync(file).toString(), this)
                else if(file.match(/.xml$/)) this[file.replace(template, '')] = parse(readFileSync(file).toString())
                else this[file.replace(template, '')] = readFileSync(file);
            })
        }
        else if(template.name == 'Workbook') {}
        else throw Error('Workbook object badly initialized\nTry creating a empty object or use "Workbook.read(workbookPath)" to read an existing Excel file.')
    }

    get style() {return this['xl/styles.xml']}
    get sharedStrings() {return this['xl/sharedStrings.xml']}
    
    static async read(path, cb=undefined) {
        let xls = new Workbook();
        let root = await JSZip.loadAsync(readFileSync(path));
        xls['xl/sharedStrings.xml'] = parse(await root.file('xl/sharedStrings.xml').async('string'));
        for(let file in root.files) {
            if(file.match(/sheet(\d){1,}.xml$/)) xls[file] = new Sheet(await root.file(file).async('string'), xls)
            else if(file.match(/.xml$/)) xls[file] = parse(await root.file(file).async('string'))
            else xls[file] = root.file(file);
        }
        return cb === undefined ? xls : cb(xls)
    }

    sheet(sheet) {
        if(typeof sheet == 'number') return this[`xl/worksheets/sheet${sheet+1}.xml`];
        let x = this['xl/workbook.xml'].find(`./sheets/sheet[@name='${sheet}']`)
        if(x === null) throw Error(`Sheet '${sheet}' not found`)
        return this[`xl/worksheets/sheet${x.attrib.sheetId}.xml`]
    }
    
    addSST(value) {
        let sst = this.sharedStrings.find('./');
        sst.attrib.uniqueCount = String((parseInt(sst.attrib.uniqueCount) || 0) + 1);
        appendElements(sst, {si: {t: {text: value}}});
    }

    save(path, cb=undefined) {
        let zip = new JSZip();
        for(let file in this) {
            if(this[file]) {
                if(file.match(/.xml$/)) zip.file(file, this[file].write({'xml_declaration': true}))
                else zip.file(file, this[file]._data)
            }
        }
        zip.generateNodeStream({type:'nodebuffer',streamFiles:true})
           .pipe(createWriteStream(path))
           .on('finish', () => cb === undefined ? ()=>{} : cb());
    }
    
}

module.exports = Workbook;