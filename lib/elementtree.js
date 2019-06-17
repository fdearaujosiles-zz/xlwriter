const et = require('elementtree');

et.ElementTree.prototype.appendElements = function (parent, obj) {
    for(let s in obj) {
        if(['attrib', 'text'].includes(s) || [false,undefined].includes(obj[s])) continue;
        let i = et.SubElement(parent, s);
        i.attrib = obj[s].attrib;
        i.text = obj[s].text;
        if(parent.attrib) if(parent.attrib.count)parent.attrib.count = String(parseInt(parent.attrib.count) + 1);
        this.appendElements(i, obj[s]);
    }
}

module.exports = et;