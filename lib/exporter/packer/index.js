/*jslint nomen: true */
/*globals module, require, __dirname */
const archiver = require('archiver');
const jsonToXml = require('./jsonToXml');
const fs = require('fs');
const path = require('path');




module.exports = function (output, officeFile, cb) {
    'use strict';

    var archive = archiver('zip');

    output.on('close', function () {
        console.log(archive.pointer() + ' total bytes');
        cb && cb(null,"done");
    });

    archive.on('error', function (err) {
        cb && cb(err)
    });

    archive.pipe(output);

    var xmlDoc = jsonToXml(officeFile.document),
        styleXml = jsonToXml(officeFile.style),
        relsXml = jsonToXml(officeFile.rels),
        coreXml = jsonToXml(officeFile.properties.core),
        head = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;


    Object.keys(officeFile.appendXML).forEach((item)=>{
       if(item === 'document'){
           xmlDoc = xmlDoc.replace(/<w:body/,officeFile.appendXML[item]+'<w:body')
       }
    });
    
    xmlDoc = xmlDoc.replace(/<!\[CDATA([\s\S]+?)\]\]>/g,($$,$1)=>$1);


    archive.bulk([
        {
            expand: true,
            cwd: officeFile.templateDir,
            src: ['**', '**/.rels']
        }
    ]);

    //archive.directory(officeFile.templateDir, false);

    officeFile.files.forEach(function(item,index){
        var ext = path.extname(item);

        if(/jpg/i.test(ext)){
            ext = ".jpeg"
        }
        archive.append(fs.createReadStream(item), {
            name: 'word/media/image'+ (index+1) + ext
        });
    });

    archive.append(head + xmlDoc, {
        name: 'word/document.xml'
    });

    archive.append(head + styleXml, {
        name: 'word/styles.xml'
    });

    archive.append(head + relsXml, {
        name: 'word/_rels/document.xml.rels'
    });

//     archive.append(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`, {
//         name: '_rels/.rels'
//     });

    archive.append(head + coreXml, {
        name: 'docProps/core.xml'
    });

    archive.finalize();
    
};