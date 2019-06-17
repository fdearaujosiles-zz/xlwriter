const Sheet = require('./Sheet');

const JSZip = require("jszip");
const walkSync = require('./walkSync');
const {parse} = require('elementtree');
const {readFileSync, createWriteStream} = require('fs');

class Workbook {
    constructor(template=undefined) {
        if(template === undefined) {
            walkSync('assets/template/').forEach(file => {
                if(file.match(/sheet(\d){1,}.xml$/)) this[file.replace('assets/template/', '')] = new Sheet(readFileSync(file).toString(), this)
                else if(file.match(/.xml$/)) this[file.replace('assets/template/', '')] = parse(readFileSync(file).toString())
                else this[file.replace('assets/template/', '')] = readFileSync(file);
            })
        }
        else if(template.name == 'Workbook') {}
        else throw Error('Workbook object badly initialized\nTry creating a empty object or use "Workbook.read(workbookPath)" to read an existing Excel file.')
    }

    get style() {return this['xl/styles.xml']}
    
    static async read(path, cb=undefined) {
        let xls = new Workbook();
        let root = await JSZip.loadAsync(readFileSync(path));
        for(let file in root.files) {
            if(file.match(/sheet(\d){1,}.xml$/)) xls[file] = new Sheet(await root.file(file).async('string'), xls)
            else if(file.match(/.xml$/)) xls[file] = parse(await root.file(file).async('string'))
            else xls[file] = root.file(file);
        }
        return cb === undefined ? xls : cb(xls)
    }

    sheet(sheetName) {
        let x = this['xl/workbook.xml'].find(`./sheets/sheet[@name='${sheetName}']`)
        if(x === null) throw Error(`Sheet '${sheetName}' not found`)
        return this[`xl/worksheets/sheet${x.attrib.sheetId}.xml`]
    }
    
    save(path, cb=undefined) {
        let zip = new JSZip();
        for(let file in this) {
            if(file.match(/.xml$/)) zip.file(file, this[file].write({'xml_declaration': true}))
            else zip.file(file, this[file]._data)
        }
        zip.generateNodeStream({type:'nodebuffer',streamFiles:true})
           .pipe(createWriteStream(path))
           .on('finish', () => cb === undefined ? ()=>{} : cb());
    }
    
}

module.exports = Workbook;