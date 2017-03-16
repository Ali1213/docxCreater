/*jslint nomen: true */
/*globals exports, require */
var border = require('./border');
var style = require('./style');
var unorderedList = require('./unorderedList');
var align = require('./alignment');
var pageBreak = require('./break');
var indent = require('./indent');
const reNodes = require('./replaceChildNodes');
const fontStyle = require("./fontStyle");
const space = require('./space');




module.exports = function (text,opt) {
    if (text === undefined) {
        text = '';
    }
    var paragraph = {
        'w:p': [{
            _attr: {
                'w:rsidR':'00CE74E5',
                'w:rsidRDefault':'0072759D'
            }
        }, {
            'w:pPr': [{
                _attr: {}
            }]
        }, {
            'w:r': [{
                'w:rPr':[{
                    _attr: {}
                },{
                    'w:rFonts':[{
                        _attr:{
                            'w:ascii':'宋体',
                            'w:eastAsia':'宋体',
                            'w:hAnsi':'宋体'
                        }
                    }]
                }]
            },{
                'w:t': text
            }]
        }],
        heading1: function () {
            var paragraphProperties = this['w:p'][1]['w:pPr'];
            paragraphProperties.push(style.style('Heading1'));
            return this;
        },
        heading2: function () {
            var paragraphProperties = this['w:p'][1]['w:pPr'];
            paragraphProperties.push(style.style('Heading2'));
            return this;
        },
        heading3: function () {
            var paragraphProperties = this['w:p'][1]['w:pPr'];
            paragraphProperties.push(style.style('Heading3'));
            return this;
        },
        title: function () {
            var paragraphProperties = this['w:p'][1]['w:pPr'];
            paragraphProperties.push(style.style('Title'));
            return this;
        },
        thematicBreak: function () {
            var paragraphProperties = paragraph['w:p'][1]['w:pPr'],
                bord = border.thematicBreak();

            paragraphProperties.push(bord);
            return this;
        },
        pageBreak: function () {
            var paragraphProperties = paragraph['w:p'][1]['w:pPr'],
                pBreak = pageBreak.pageBreak();

            paragraphProperties.push(pBreak);
            return this;
        },
        addTabStop: function (tab) {
            var paragraphProperties = paragraph['w:p'][1]['w:pPr'];

            paragraphProperties.push(tab);
            return this;
        },
        bullet: function () {
            var paragraphProperties = paragraph['w:p'][1]['w:pPr'];

            paragraphProperties.push(unorderedList.bullet());
            paragraphProperties.push(unorderedList.numberProperties());
            return this;
        },
        /*下面是修改后的代码*/
        alignment:function(str){
            if(!["right","left","center","justified"].includes(str)){
                str = "left";
            }
            let paragraphProperties = this['w:p'][1]['w:pPr'];
            paragraphProperties.push(align[str]());
            return this;
        },
        addIndent:function(num){
            let paragraphProperties = this['w:p'][1]['w:pPr'];
            reNodes(paragraphProperties,"w:ind",indent(num));
            return this;
        },
        addText: function (run) {
            var paragraphs = this['w:p'];
            paragraphs.push(run);
            return this;
        },
        pFontStyle:function({
            fontSize = 12
        }={}){
            let paragraphProperties = this['w:p'][1]['w:pPr'];
            let p = fontStyle().fontSize(fontSize);
            reNodes(paragraphProperties,"w:rPr",p);

            for(let elem of this['w:p']){
                if(elem['w:r']){
                    reNodes(elem['w:r'],"w:rPr",p);
                }
            }

            return this;
        },
        setSpace:function(opt){
            
/*            if(typeof opt['after'])
            
            var obj = {};
            obj['after']= after * 100;

            obj['before']= before * 100;
            obj['line']= line * 20;
            obj['lineType']= lineType;
            {
                after = 0.5,
                    before = 0,
                    line = 8,
                    lineType = "auto"
            }={}*/

            let paragraphProperties = this['w:p'][1]['w:pPr'];
            reNodes(paragraphProperties,'w:spacing',space(opt))
            return this;
        }
    };

/*    if(opt.indent){
        paragraph['w:p'][1]['w:pPr'].push(indent(num));
    }*/

    return paragraph;
};
