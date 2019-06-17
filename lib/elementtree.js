const et = require('elementtree');

et.appendElements = function (parent, obj, count=false) {
    for(let s in obj) {
        if(['attrib', 'text'].includes(s) || [false,undefined].includes(obj[s])) continue;
        let i = et.SubElement(parent, s, obj[s].attrib);
        i.text = obj[s].text;
        if(!parent.attrib) parent.attrib = {} 
        if(parent.attrib.count || count) parent.attrib.count = String((parseInt(parent.attrib.count) || 0) + 1);
        et.appendElements(i, obj[s]);
    }
}

module.exports = et;