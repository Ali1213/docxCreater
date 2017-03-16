/*jslint nomen: true */
/*globals exports, require */
const row = require('./row');
const cell = require('./cell');
const reChild = require('../replaceChildNodes');




const createGridCol = function(columnCount){

    if(!columnCount){
        columnCount = 3
    }
    let columns = [],
        width = parseInt(9015 / columnCount, 10),
        i=0;


    while(i<columnCount){
        columns.push({
            'w:gridCol': [{
                _attr: {
                    'w:w': width
                }
            }]
        });
        i++;
    }

    return columns;
};

exports.createTable = function (columnCount) {
    'use strict';

    var table = {
        'w:tbl': [{
            _attr: {}
        }, {
            'w:tblPr': [{
                "w:tblStyle":[{
                    _attr:{
                        "w:val": "TableGrid"
                    }
                }]
            },{
                "w:tblW":[{
                    _attr:{
                        "w:w": "0",
                        "w:type":"auto"
                    }
                }]
            },{
               'w:tblBorders':[{
                   _attr:{}
               },{
                   'w:top':[{
                       _attr:{
                           'w:val': 'single',
                           'w:sz':'1',
                           'w:space':'0',
                           'w:color':'000000'
                       }
                   }]
               },{
                   'w:left':[{
                       _attr:{
                           'w:val': 'single',
                           'w:sz':'1',
                           'w:space':'0',
                           'w:color':'000000'
                       }
                   }]
               },{
                   'w:bottom':[{
                       _attr:{
                           'w:val': 'single',
                           'w:sz':'1',
                           'w:space':'0',
                           'w:color':'000000'
                       }
                   }]
               },{
                   'w:right':[{
                       _attr:{
                           'w:val': 'single',
                           'w:sz':'1',
                           'w:space':'0',
                           'w:color':'000000'
                       }
                   }]
               },{
                   'w:insideH':[{
                       _attr:{
                           'w:val': 'single',
                           'w:sz':'1',
                           'w:space':'0',
                           'w:color':'000000'
                       }
                   }]
               },{
                   'w:insideV':[{
                       _attr:{
                           'w:val': 'single',
                           'w:sz':'1',
                           'w:space':'0',
                           'w:color':'000000'
                       }
                   }]
               }]
            },{
                "w:jc":[{
                    _attr:{
                        "w:val": "center"
                    }
                }]
            },{
                "w:tblLook":[{
                    _attr:{
                        "w:val": "04A0",
                        "w:firstRow": "1",
                        "w:lastRow": "0",
                        "w:firstColumn": "1",
                        "w:lastColumn": "0",
                        "w:noHBand": "0",
                        "w:noVBand": "1"
                    }
                }]
            }]
        }, 
            {
            'w:tblGrid': createGridCol(columnCount)
        }],
        addRow: function (row) {
               this["w:tbl"].push(row);
            return this;
        }
        /*新增的方法*/,
        setColWidth:function(){
            //还没有写完
            return this;
        },
        resetTalGrid:function(num){
            reChild(this['w:tbl'],'w:tblGrid',{
                'w:tblGrid':createGridCol(num)
            });
            return this;
        }
    };
    return table;
};

exports.createRow = row.createRow;
exports.createCell = cell.createCell;