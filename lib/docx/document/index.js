


const getPgSeting = function(obj, n, tag){
    let body = obj['w:document'][1]['w:body'];
    return body[body.length-1]['w:sectPr'][n][tag][0]['_attr'];
};

const changePageWH = function(pzObj,isLandscope){
    // isLandscope为true 时 w>h 横向
    let [w,h] = [Number(pzObj['w:w']),Number(pzObj['w:h'])];
    if(isLandscope !== (w>h)){
            [pzObj['w:w'],pzObj['w:h']] = [h,w];
    }
};



module.exports = function(){
    let doc = {
        'w:document': [{
            _attr: {
                'xmlns:wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
                'xmlns:cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
                'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'xmlns:o': 'urn:schemas-microsoft-com:office:office',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'xmlns:m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                'xmlns:v': 'urn:schemas-microsoft-com:vml',
                'xmlns:wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
                'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'xmlns:w10': 'urn:schemas-microsoft-com:office:word',
                'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
                'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
                'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
                'xmlns:wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
                'xmlns:wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
                'xmlns:wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
                'xmlns:wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
                'mc:Ignorable': 'w14 w15 wp14'
            }
        }, {
            'w:body': [{
                'w:sectPr': [{
                    _attr: {
                        'w:rsidR': '00CE74E5'
                    }
                }, {
                    'w:pgSz': [{
                        _attr: {
                            'w:w': '11906',
                            'w:h': '16838'
                        }
                    }]
                }, {
                    'w:pgMar': [{
                        _attr: {
                            'w:top': '1440',
                            'w:right': '1800',
                            'w:bottom': '1440',
                            'w:left': '1800',
                            'w:header': '851',
                            'w:footer': '992',
                            'w:gutter': '0'
                        }
                    }]
                }, {
                    'w:cols': [{
                        _attr: {
                            'w:space': '425'
                        }
                    }]
                }, {
                    'w:docGrid': [{
                        _attr: {
                            'w:type': 'lines',
                            'w:linePitch':'312'
                        }
                    }]
                }]
            }]
        }],
        setPageWidth:function(num){
            let sz = getPgSeting(this,1,'w:pgSz');
            sz['w:w'] = num;
            return this;
        },
        setPageHeight:function(num){
            let sz = getPgSeting(this,1,'w:pgSz');
            sz['w:h'] = num;
            return this;
        },
        setOrient:function(bol){
            let sz = getPgSeting(this,1,'w:pgSz');
            if(bol){
                sz['w:orient'] = 'landscape';
            }else{
                delete sz['w:orient']
            }
            changePageWH(sz,bol);
            return this;
        },
        setCode:function(num){
            let sz = getPgSeting(this,1,'w:pgSz');
            sz['w:code'] = num;
            return this;
        },
        setCol:function(num){
            let sz = getPgSeting(this,3,'w:cols');
            sz['w:num'] = num;
            return this;
        },
        setSpace:function(num){

            let sz = getPgSeting(this,3,'w:cols');
            sz['w:space'] = num;
            return this;
        },
        setMg:function(tagname,value){

            let sz = getPgSeting(this,2,'w:pgMar');
            sz[tagname] = value;
            return this;
        }
    };

    return doc
}