const path = require('path');
/*
* inset 嵌入 默认类型；
* */

var createRels = function (id, type, Target) {
    'use strict';

    return {
        "Relationship": {
            _attr: {
                'Id': id,
                'Type': type,
                'Target': Target
            }
        }
    };
};



var reLayout = function(style,wrap,str){
    var arr = style.split(";");
    var json = {};
    var newStr = "";
    arr.forEach((item,index)=>{
        if(!item.trim()){
            return ;
        }
        let arr2 = item.split(':');
        json[arr2[0]] = arr2[1];
    })

    switch(str){
        case "inset":
            break;
        case "square":
            json["position"] = "absolute";
            // json["left"] = 0;
            // json["text-align"] = "left";
            // json["margin-left"] = "21pt";
            // json["margin-top"] = "2.45pt";
            wrap.push({
                "w10:wrap":[{
                    _attr:{
                        "type":"square"
                    }
                }]
            });
            break;
        case "tight":
            json["position"] = "absolute";
            // json["left"] = 0;
            // json["text-align"] = "left";
            // json["margin-left"] = "21pt";
            // json["margin-top"] = "2.45pt";
            json["z-index"] = "-251655168";
            // json["margin-top"] = "2.45pt";

            wrap.push({
                "w10:wrap":[{
                    _attr:{
                        "type":"tight"
                    }
                }]
            });
            break;
        case "through":
            json["position"] = "absolute";
            // json["left"] = 0;
            // json["text-align"] = "left";
            // json["margin-left"] = "21pt";
            // json["margin-top"] = "2.45pt";
            json["z-index"] = "-251655168";
            // json["margin-top"] = "2.45pt";
            wrap.push({
                "w10:wrap":[{
                    _attr:{
                        "type":"through"
                    }
                }]
            });
            break;
        case "topAndBottom":
            json["position"] = "absolute";
            // json["left"] = 0;
            // json["text-align"] = "left";
            // json["margin-left"] = "21pt";
            // json["margin-top"] = "2.45pt";
            json["z-index"] = "-251655168";
            // json["margin-top"] = "2.45pt";
            wrap.push({
                "w10:wrap":[{
                    _attr:{
                        "type":"topAndBottom"
                    }
                }]
            });
            break;
        case "underText":
            json["position"] = "absolute";
            // json["left"] = 0;
            // json["text-align"] = "left";
            // json["margin-left"] = "21pt";
            // json["margin-top"] = "2.45pt";
            json["z-index"] = "-251657216";
            // json["margin-top"] = "2.45pt";
            break;
        case "aboveText":
            json["position"] = "absolute";
            // json["left"] = 0;
            // json["text-align"] = "left";
            // json["margin-left"] = "21pt";
            // json["margin-top"] = "2.45pt";
            json["z-index"] = "251661312";
            // json["margin-top"] = "2.45pt";
            break;
    }

    Object.keys(json).forEach((item)=>{
        newStr+=`${item}:${json[item]};`
    });
    return newStr.slice(0,-1);
};

//创建link需要做下面的事情
//1.在document.xml创建链接;
//2.在_rels/document.xml.rels创建标签
module.exports = function (filepath, width, height,rels,files) {

    var num = rels['Relationships'].length;
    var id = 'rId' + num;
    var type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
/*
    var shape = {
        "x": 450,
        "y": parseInt(450 /width  * height)
    };
*/
    // 如果图片宽度px换算成pt大于页面宽度450pt
    let x,y;
    if(width*4/3>450){
        x=450;
        y=parseInt(450 /width  * height);
    }else{
        [x,y]=[width*4/3,height*4/3];
    }


    files.push(filepath);

    let ext = path.extname(filepath);
    if(/jpg/i.test(ext)){
        ext = ".jpeg"
    }
    var target = "media/image" + files.length + ext;

    var linkDoc = {
        "w:r":[{
           "w:pict":[{
               _attr:{}
           },{
               "v:shape":[{
                   _attr:{
                       "id":"_x0000_i1025",
                       "style":`width:${x}pt;height:${y}pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-relative:page;mso-height-relative:page`,
                       "coordsize":"",
                       "o:spt":"100",
                       "adj":"0,,0",
                       "path":"",
                       "stroked":"f"
                   }
               },{
                   "v:stroke":[{
                       _attr:{
                           "joinstyle":"miter"
                       }
                   }]
               },{
                   "v:imagedata":[{
                       _attr:{
                           "r:id":id,
                           "o:title":"3"
                       }
                   }]
               },{
                   "v:formulas":[{}]
               },{
                   "v:path":[{
                       _attr:{
                           "o:connecttype":"segments"
                       }
                   }]
               }]
           }] 
        }],
        "layout":function(str){
            if(!["inset","square","tight","through","topAndBottom","underText","aboveText"].includes(str)){
                str = "inset";
            }
            let style = this["w:r"][0]["w:pict"][1]["v:shape"][0]["_attr"];
            let wrap = this["w:r"][0]["w:pict"][1]["v:shape"];

            style.style = reLayout(style.style,wrap,str);
            return this;
        }
    }
 /*   var linkDoc = {
        'w:r': [{
            _attr: {}
        }, {
            'w:drawing': [{
                _attr: {}
            }, {
                'wp:inline': [{
                    _attr: {
                        "distT": "0",
                        "distB": "0",
                        "distL": "0",
                        "distR": "0"
                    }
                }, {
                    'wp:extent': [{
                        _attr: {
                            'cx': shape.x,
                            'cy': shape.y
                        }
                    }]
                }, /!*{
                 'wp:effectExtent':[{
                 _attr:{
                 'l':"0",
                 't':"0",
                 'r':"2540",
                 'b':"2540"
                 }
                 }]
                 },*!/{
                    'wp:docPr': [{
                        _attr: {
                            'id': "" + num,
                            'name': "图片 " + num,
                            "descr": ""
                        }
                    }]
                }, {
                    'wp:cNvGraphicFramePr': [{
                        _attr: {}
                    }, {
                        "a:graphicFrameLocks": [{
                            _attr: {
                                "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                                "noChangeAspect": "1"
                            }
                        }]
                    }]
                }, {
                    "a:graphic": [{
                        _attr: {
                            'xmlns:a': "http://schemas.openxmlformats.org/drawingml/2006/main",
                            'cy': shape.y
                        }
                    }, {
                        "a:graphicData": [{
                            _attr: {
                                'uri': "http://schemas.openxmlformats.org/drawingml/2006/picture"
                            }
                        }, {
                            "pic:pic": [{
                                _attr: {
                                    "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                }
                            }, {
                                "pic:nvPicPr": [{
                                    "pic:cNvPr": [{
                                        _attr: {
                                            "id": "" + num,
                                            "name": "Picture " + num,
                                            "descr": ""
                                        }
                                    }]
                                }, {
                                    "pic:cNvPicPr": [{
                                        "a:picLocks": [{
                                            _attr: {
                                                "noChangeAspect": "1",
                                                "noChangeArrowheads": "1"
                                            }
                                        }]
                                    }]
                                }]
                            }, {
                                "pic:blipFill": [{
                                    "a:blip": [{
                                        _attr: {
                                            "r:embed": id
                                        }
                                    }/!*,{
                                     "a:extLst":[{
                                     "a:ext":[{
                                     _attr:{
                                     "uri":"{28A0092B-C50C-407E-A947-70E740481C1C}"
                                     }
                                     },{
                                     "a14:useLocalDpi":[{
                                     _attr:{
                                     "xmlns:a14":"http://schemas.microsoft.com/office/drawing/2010/main",
                                     "val":"0"
                                     }
                                     }]
                                     }]
                                     }]
                                     }]
                                     }*!/, {
                                        "a:stretch": [{
                                            "a:fillRect": [{}]
                                        }]
                                    }]
                                }, {
                                    "pic:spPr": [{
                                        _attr: {
                                            "bwMode": "auto"
                                        }
                                    }, {
                                        "a:xfrm": [{
                                            "a:off": [{
                                                _attr: {
                                                    "x": "0",
                                                    "y": "0"
                                                }
                                            }]
                                        }, {
                                            "a:ext": [{
                                                _attr: {
                                                    "cx": shape.x,
                                                    "cy": shape.y
                                                }
                                            }]
                                        }]
                                    }, {
                                        "a:prstGeom": [{
                                            _attr: {
                                                "prst": "rect"
                                            }
                                        }, {
                                            "a:avLst": [{}]
                                        }]
                                    }]
                                }]
                            }]
                        }]
                    }]
                }]
            }
        ]
    }]}*/
/*    var setStyle = function(json,value1,value2,value3,value4,value5,value6){
        json["mso-position-horizontal"] = value1;
        json["mso-position-horizontal-relative"] = value2;
        json["mso-position-vertical"] = value3;
        json["mso-position-vertical-relative"] = value4;
        json["mso-width-relative"] = value5;
        json["mso-height-relative"] = value6;
    };*/


    var linkRels = rels.addRels(id, type, target);
    var bookmarkStart = {
        "w:bookmarkStart": [{
            _attr: {
                'w:id': num,
                'w:name': '_GoBack'
            }
        }]
    };
    var bookmarkEnd = {
        "w:bookmarkEnd": [{
            _attr: {
                'w:id': num
            }
        }]
    };

    rels['Relationships'].push(linkRels);
    return linkDoc;
}
