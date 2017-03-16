/*jslint nomen: true */
/*globals exports, module, require */

function createLsdException(name, uiPriority, qFormat, semiHidden, unhideWhenUsed) {
    'use strict';

    return [{
        _attr: {
            'w:name': name,
            'w:uiPriority': uiPriority,
            'w:qFormat': qFormat,
            'w:semiHidden': semiHidden,
            'w:unhideWhenUsed': unhideWhenUsed
        }
    }];
}

module.exports = function () {
    'use strict';

    let style = {
        'w:styles': [{
            _attr: {
                'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
                'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
                'xmlns:w16se':'http://schemas.microsoft.com/office/word/2015/wordml/symex',
                'mc:Ignorable': 'w14 w15 w16se'
            }
        }, {
            'w:docDefaults': [{
                'w:rPrDefault': [{
                    'w:rPr': [{
                        'w:rFonts': [{
                            _attr: {
                                'w:asciiTheme': "minorHAnsi",
                                'w:eastAsiaTheme': "minorHAnsi",
                                'w:hAnsiTheme': "minorHAnsi",
                                'w:cstheme': "minorBidi"
                            }
                        }]
                    }, {
                        'w:kern': [{
                            _attr: {
                                'w:val': "2"
                            }
                        }]
                    },{
                        'w:sz': [{
                            _attr: {
                                'w:val': "21"
                            }
                        }]
                    }, {
                        'w:szCs': [{
                            _attr: {
                                'w:val': "22"
                            }
                        }]
                    }, {
                        'w:lang': [{
                            _attr: {
                                'w:val': "en-US",
                                'w:eastAsia': "zh-CN",
                                'w:bidi': "ar-SA"
                            }
                        }]
                    }]
                }]
            }, {
                'w:pPrDefault': [{
                    'w:pPr': [{
                        'w:spacing': [{
                            _attr: {
                                'w:after': "160",
                                'w:line': "259",
                                'w:lineRule': "auto"
                            }
                        }]
                    }]
                }]
            }]
        }, {
            'w:latentStyles': [{
                _attr: {
                    'w:defLockedState': "0",
                    'w:defUIPriority': "99",
                    'w:defSemiHidden': "0",
                    'w:defUnhideWhenUsed': "0",
                    'w:defQFormat': "0",
                    'w:count': "371"
                }
            }, {
                'w:lsdException': createLsdException('Normal', 0, 1)
            }, {
                'w:lsdException': createLsdException("heading 1", 9, 1)
            }, {
                'w:lsdException': createLsdException("heading 2", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 3", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 4", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 5", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 6", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 7", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 8", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("heading 9", 9, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 1", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 5", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 6", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 7", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 8", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index 9", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 1", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 2", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 3", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 4", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 5", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 6", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 7", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 8", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toc 9", 39, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Normal Indent", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("footnote text", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("annotation text", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("header", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("footer", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("index heading", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("caption", 35, 1, 1, 1)
            }, {
                'w:lsdException': createLsdException("table of figures", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("envelope address", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("envelope return", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("footnote reference", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("annotation reference", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("line number", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("page number", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("endnote reference", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("endnote text", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("table of authorities", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("macro", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("toa heading", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Bullet", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Number", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List 5", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Bullet 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Bullet 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Bullet 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Bullet 5", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Number 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Number 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Number 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Number 5", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Title", 10, 1, undefined, undefined)
            }, {
                'w:lsdException': createLsdException("Closing", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Signature", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Default Paragraph Font", 1, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text Indent", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Continue", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Continue 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Continue 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Continue 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("List Continue 5", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Message Header", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Subtitle", 11, 1, undefined, undefined)
            }, {
                'w:lsdException': createLsdException("Salutation", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Date", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text First Indent", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text First Indent 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Note Heading", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text Indent 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Body Text Indent 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Block Text", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Hyperlink", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("FollowedHyperlink", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Strong", 22, 1)
            }, {
                'w:lsdException': createLsdException("Emphasis", 20, 1)
            }, {
                'w:lsdException': createLsdException("Document Map", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Plain Text", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("E-mail Signature", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Top of Form", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Bottom of Form", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Normal (Web)", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Acronym", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Address", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Cite", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Code", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Definition", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Keyboard", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Preformatted", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Sample", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Typewriter", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("HTML Variable", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Normal Table", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("annotation subject", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("No List", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Outline List 1", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Outline List 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Outline List 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Simple 1", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Simple 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Simple 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Classic 1", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Classic 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Classic 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Classic 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Colorful 1", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Colorful 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Colorful 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Columns 1", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Columns 2", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Columns 3", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Columns 4", undefined, undefined, 1, 1)
            }, {
                'w:lsdException': createLsdException("Table Columns 5", undefined, undefined, 1, 1)
            }]
        },{
            'w:style':[{
                _attr : {
                    'w:type':'paragraph',
                    'w:default':'1',
                    'w:styleId':'a'
                }
            },{
                'w:name':[{
                    _attr:{
                        'w:val':'Normal'
                    }
                }]
            },{
                'w:qFormat':[{
                    _attr:{}
                }]
            },{
                'w:pPr':[{
                    _attr:{}
                },{
                    'w:widowControl':[{
                        _attr:{
                           'w:val':'0' 
                        }
                    }]
                },{
                    'w:jc':[{
                        _attr:{
                            'w:val':'both'
                        }
                    }]
                }]
            }]
        },{
            'w:style':[{
                _attr : {
                    'w:type':'character',
                    'w:default':'1',
                    'w:styleId':'a0'
                }
            },{
                'w:name':[{
                    _attr:{
                        'w:val':'Default Paragraph Font'
                    }
                }]
            },{
                'w:uiPriority':[{
                    _attr:{
                        'w:val':'1'
                    }
                }]
            },{
                'w:semiHidden':[{
                    _attr:{}
                }]
            },{
                'w:unhideWhenUsed':[{
                    _attr:{}
                }]
            }]
        },{
            'w:style':[{
                _attr : {
                    'w:type':'table',
                    'w:default':'1',
                    'w:styleId':'a1'
                }
            },{
                'w:name':[{
                    _attr:{
                        'w:val':'Normal Table'
                    }
                }]
            },{
                'w:uiPriority':[{
                    _attr:{
                        'w:val':'99'
                    }
                }]
            },{
                'w:semiHidden':[{
                    _attr:{}
                }]
            },{
                'w:unhideWhenUsed':[{
                    _attr:{}
                }]
            },{
                'w:tblPr':[{
                    _attr:{}
                },{
                    'w:tblInd':[{
                        _attr:{
                            'w:w':'0',
                            'w:type':'dxa'
                        }
                    }]
                },{
                    'w:tblCellMar':[{
                        _attr:{}
                    },{
                        'w:top':[{
                            _attr:{
                                'w:w':'0',
                                'w:type':'dxa'
                            }
                        }]
                    },{
                        'w:left':[{
                            _attr:{
                                'w:w':'108',
                                'w:type':'dxa'
                            }
                        }]
                    },{
                        'w:bottom':[{
                            _attr:{
                                'w:w':'0',
                                'w:type':'dxa'
                            }
                        }]
                    },{
                        'w:right':[{
                            _attr:{
                                'w:w':'108',
                                'w:type':'dxa'
                            }
                        }]
                    }]
                }]
            }]
        },{
            'w:style':[{
                _attr : {
                    'w:type':'numbering',
                    'w:default':'1',
                    'w:styleId':'a2'
                }
            },{
                'w:name':[{
                    _attr:{
                        'w:val':'No List'
                    }
                }]
            },{
                'w:uiPriority':[{
                    _attr:{
                        'w:val':'99'
                    }
                }]
            },{
                'w:semiHidden':[{
                    _attr:{}
                }]
            },{
                'w:unhideWhenUsed':[{
                    _attr:{}
                }]
            }]
        },{
            'w:style':[{
                _attr : {
                    'w:type':'character',
                    'w:styleId':'a3'
                }
            },{
                'w:name':[{
                    _attr:{
                        'w:val':'Hyperlink'
                    }
                }]
            },{
                'w:basedOn':[{
                    _attr:{
                        'w:val':'a0'
                    }
                }]
            },{
                'w:uiPriority':[{
                    _attr:{
                        'w:val':'99'
                    }
                }]
            },{
                'w:unhideWhenUsed':[{
                    _attr:{}
                }]
            },{
                'w:rsid':[{
                    _attr:{
                        'w:val':'00041FF5'
                    }
                }]
            },{
                'w:rPr':[{
                    _attr:{}
                },{
                    'w:color':[{
                        _attr:{
                            'w:val':'0563C1',
                            'w:themeColor':'hyperlink'
                        }
                    }]
                },{
                    'w:u':[{
                        _attr:{
                            'w:val':'single'
                        }
                    }]
                }]
            }]
        }]
    };

    return style;
};