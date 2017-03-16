/*globals exports */
/*jslint nomen: true */

const vMergeStart = {
    'w:vMerge':[{
        _attr:{
           'w:val':'restart'
        }
    }]
};

const vMergeOn = {
    'w:vMerge':[{
        _attr:{}
    }]
};

const gridSpan = (num)=>{
    return {
        'w:gridSpan':[{
            _attr:{
                'w:val':num
            }
        }]
    }
};

exports.createCell = function () {
    'use strict';

    var cell = {
        'w:tc': [{
            _attr: {

            }
        }, {
            'w:tcPr': [{
                'w:tcW': [{
                    _attr: {
                        'w:w': '3006',
                        'w:type': 'dxa'
                    }
                }]
            }]
        }],
        addParagraph: function (paragraph) {
            this['w:tc'].push(paragraph);
            return this;
        },
        addTable: function (table) {
            return this;
        }
        /*新增的方法*/,
        addRowSpanStart:function(){
            this['w:tc'][1]['w:tcPr'].push(vMergeStart)
            return this
        },
        addRowSpanOn:function(){
            this['w:tc'][1]['w:tcPr'].push(vMergeOn)
            return this
        },
        addColSpan:function(num){
            this['w:tc'][1]['w:tcPr'].push(gridSpan(num));
            return this
        }

    };
    return cell;
};