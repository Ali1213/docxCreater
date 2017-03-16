
var createRels = function (id, type, Target) {
    'use strict';

    return {
        "Relationship":{
            _attr: {
                'Id': id,
                'Type': type,
                'Target': Target,
                "TargetMode":"External"
            }
        }
    };
}


//创建link需要做下面的事情
//1.在document.xml创建链接；
//2.在_rels/document.xml.rels创建标签
module.exports = function(text,link,rels){


    var num = rels['Relationships'].length;
    var id = 'rId' + num;
    var type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    var linkDoc = {
        'w:hyperlink': [{
            _attr: {
                'r:id': id,
                'w:history':'1'
            }
        },{
            'w:r': [{
                'w:rPr': [{
                    _attr: {
                    }
                },{
                    'w:rStyle':[{
                        _attr:{
                           'w:val':"a3"
                        }
                    }]
                }]
            }, {
                'w:t': text
            }]
        }]
    };

    var linkRels = rels.addRels(id,type,link,"External");
    var bookmarkStart = {
        "w:bookmarkStart":[{
            _attr: {
                'w:id': num,
                'w:name':'_GoBack'
            }
        }]
    };
    var bookmarkEnd = {
        "w:bookmarkEnd":[{
            _attr: {
                'w:id': num
            }
        }]
    }

    rels['Relationships'].splice(rels['Relationships'].length-1,0,linkRels);

    return linkDoc;
};
