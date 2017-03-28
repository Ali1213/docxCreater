const document = require("./lib/docx/document/index");
const style = require("./lib/docx/style");
const rels = require("./lib/docx/rels");
const paragraph = require('./lib/docx/paragraph');
const text = require('./lib/docx/text');
const packer = require('./lib/exporter');
const properties = require('./lib/docx/properties');
const links = require('./lib/docx/links');
const images = require('./lib/docx/images');
const table = require('./lib/docx/table');
const path = require("path");
const fs = require("fs");


class Docx {
    constructor({
        filePath = path.join(__dirname,"1.docx"),
        defaultFontSize = 12
    }={}) {

        //content
        this.document = document();
        this.style = style();
        this.rels = rels();

        //files
        this.files = [];

        // 内部存储
        this.content = [];

        this.last = null;


        //
        this.appendXML = {};

        //constant
        this.constant = {
            "extention":'docx',
            "templateDir":path.join(__dirname,'template')
        };
        this.output = filePath;
        this.defaultFontSize = defaultFontSize;

        //字号大小
        this.fontSizeType = {
            "初号":42,
            "小初":36,
            "一号":26,
            "小一":24,
            "二号":22,
            "小二":18,
            "三号":16,
            "小三":15,
            "四号":14,
            "小四":12,
            "五号":10.5,
            "小五":9,
            "六号":7.5,
            "小六":6.5,
            "七号":5.5,
            "八号":5
        };
        //页面大小
        this.pageSize = {
            'A4':[11906,16838],
            'A3':[16838,23811]
        }

    }
    //下面统一处理段落样式
    /*
     * indent:段首缩进字符，默认是0个字符
     * alignment:段落对齐方式，"right","left","center","justified",4选1，默认left
     * */
    addParagraph(str,{
        indent = 0,
        alignment = 'left'
    }={}){
        this.last = paragraph(str).addIndent(indent).alignment(alignment);

        if(this.paragraphSpace){
            Docx.setSpace(this.last,this.paragraphSpace)
        }

        this.content.push(this.last);
        return this;
    }

    addIndent(indent){
        if(typeof indent !== 'number'){
            return this;
        }
        this.last.addIndent(indent);
        return this;
    }

    //统一设置段落字体大小
    rStyle({
        fontSize = 12
    }={}){
        this.last.pFontStyle({
            fontSize:Docx.convertFontSize(fontSize,this.fontSizeType)
        });
        return this;
    }

    //下面统一处理文字样式
    //fontSize : number/string 字体大小
    //subScript : boolean 是否是下标
    //superScript : boolean 是否是下标
    //bold : boolean 是否是黑体
    //italic : boolean 是否是下标
    //underline : boolean 是否是下标
    //strike : boolean 是否是下标
    addText(str,{
        fontSize = 12,
        subScript = false,
        superScript = false,
        bold = false,
        italic = false,
        underline = false,
        strike = false,
    }={}){
        let obj;

        if(typeof str === 'string'){
            obj = text(str);
        }else{
            obj = str;
        }


        if(typeof fontSize !== 'number'){
            fontSize = this['fontSizeType'][fontSize] ? this['fontSizeType'][fontSize] : 12;
        }
        obj.fontSize(fontSize);


        if(subScript){
            obj.subScript();
        }

        if(superScript){
            obj.superScript();
        }

        if(bold){
            obj.bold();
        }

        if(italic){
            obj.italic();
        }

        if(underline){
            obj.underline();
        }

        if(strike){
            obj.strike();
        }

        this.last.addText( obj );
        return this;
    }


    /*
     * 页面样式配置
     * 1参为w:pgSz的设置，目前已知为页面大小，页面方向
     *
     *
     * type优先,目前是A3,A4
     *
     * */

    setPage({
        type = "",
        width = 0,
        height = 0,
        orient = false,
        cols = 1,
        colsSpace = 0,
        pageTop = 0,
        pageRight = 0,
        pageBottom = 0,
        pageLeft = 0,
        header = 0,
        footer = 0,
        gutter = 0
    }){

        if(width){
            this.document.setPageWidth(width);
        }

        if(height){
            this.document.setPageHeight(height);
        }

        if( this.pageSize[type] ){
            this.document.setPageWidth(this.pageSize[type][0]).setPageHeight(this.pageSize[type][1]);

            //下一条在A3时被添加，不知道有什么用
            this.document.setCode(8)
        }

        if(orient){
            this.document.setOrient(true)
        }


        //暂时限制3
        if(cols>1 && cols<4){
            this.document.setCol(2);

            if(colsSpace){
                this.document.setSpace(colsSpace);
            }
        }


        if(pageTop){
            this.document.setMg('w:top',pageTop)
        }
        if(pageRight){
            this.document.setMg('w:right',pageRight)
        }
        if(pageBottom){
            this.document.setMg('w:bottom',pageBottom)
        }
        if(pageLeft){
            this.document.setMg('w:left',pageLeft)
        }
        if(header){
            this.document.setMg('w:header',header)
        }
        if(footer){
            this.document.setMg('w:footer',footer)
        }
        if(gutter){
            this.document.setMg('w:gutter',gutter)
        }


        return this;
    }




    //
    addLink(text,link){
        this.last.addText(links(text,link,this.rels));
        return this;
    }

    setPSpace(opt){
        if(typeof opt !== 'object'){
            return this;
        }

        this.paragraphSpace = opt;

        if(this.last){
            Docx.setSpace(this.last,opt)
        }

        return this;
    }

    // 做一件很污的事，允许你随意插入任何类型的对象
    insertDocument(xml){

        if(typeof xml === 'string'){
            if(!this.last){
                this.addParagraph();
            }

            this.last.addText(Docx.convertXML2XMLObj(xml));
        }



        if(xml['w:p'] || xml['w:tbl']){
            this.content.push(xml);
        }
        return this;
    }

    //做一件更污的事情，允许往document.body最前面 插入xml
    insertDocumentXML(str){
        this.appendXML['document'] = str;
        return this;
    }


    //处理图片样式
    /*
     * type暂时因为位置的原因无法处理；
     * inset:嵌入
     * square:四周
     * tight:紧密环绕
     * through:穿越型环绕
     * topAndBottom:上下型环绕
     * underText:衬于文字下方
     * aboveText:浮于文字上方；
     * */
    addImage(filepath,width,height,type = "inset"){
        this.last.addText(images(filepath,width,height,this.rels,this.files).layout(type));
        return this;
    }

    /*
     * 1.创建表格
     * 2.处理基本的样式
     * 3.设定行和列的数量
     * 4.单元格合并的设定
     * */
    /*
     *
     * */
    addTable({
        colNum = 3,
        rowNom = 3,
        colWidth = [],
        colspan = {},
        rowspan = {},
        HTMLTable = ''
    }={}){

        if(HTMLTable){
            let tb = Docx.converHTMLTable(HTMLTable);
            this.content.push( tb );
            this.last = tb;
            return this;
        }
    }
    /*    createRow(){
     return table.createRow();
     }

     createCell(){
     return table.createCell();
     }*/

    createMaxRightTabStop(){
        return Docx.createTabTemplate('right', 'RIGHT_MARGIN');
    }

    createRightTabStop(){

        return Docx.createTabTemplate('right', position);
    }

    createLeftTabStop(){

        return Docx.createTabTemplate('left', position);
    }

    createCenterTabStop(){
        return Docx.createTabTemplate('center', position);
    }


    render(options,cb){
        let my = this;

        let body = my.document['w:document'][1]['w:body'];

        //我TM也不知道这个有什么用
        my.content = my.content.map((item,i)=>{
            return JSON.parse(JSON.stringify(item).replace('RIGHT_MARGIN','9026'));
        });
        //将所有段落插入文本中
        body.splice(body.length - 1, 0,...my.content);

        let officeFile = {
            "document":my.document,
            "style":my.style,
            "rels":my.rels,
            properties: {
                core: properties.core(options)
            },
            templateDir: my.constant.templateDir,
            extention: my.constant.extention,
            files:my.files,
            appendXML:this.appendXML
        };
        packer(fs.createWriteStream(my.output),officeFile,cb);
    }


    /*静态方法*/

    static createParapragh(str){
        return paragraph(str);
    }

    static createTabTemplate(value,position){
        return {
            'w:tabs': [{
                'w:tab': [{
                    _attr: {
                        'w:val': value,
                        'w:pos': position
                    }
                }]
            }]
        };
    }


    //如果传过来的fontSize是非数字
    static convertFontSize(fontSize,ft){
        if(typeof fontSize !== 'number'){
            fontSize = ft[fontSize] ? ft[fontSize] :this.defaultFontSize;
        }
        return fontSize;
    };
    /*    static likeArray2Array(obj){
     let arr = [];
     for(let i=0;i<obj.length;i++){
     arr.push(obj[i]);
     }
     return arr;
     };*/


    static converHTMLTable (HTMLTable){
        let trRe =/<tr[\s\S]+?<\/tr>/ig;
        let tdRe = /<t[dh][\s\S]+?<\/t[dh]>/ig;

        if(typeof HTMLTable !== 'string'){
            throw new Error("HTMLTable type isn't not string");
        }

        //确定table的列数量
        let colMaxNum = 0;
        let x=0,y=0;
        let rowspanList = {};
        let tb = table.createTable(0);

        HTMLTable.match(trRe).forEach((item,index)=>{
            console.log(index);
            let tdArr = item.match(tdRe);
            let trObj = table.createRow();
            let cellArr = [];
            if(tdArr.length > colMaxNum){
                colMaxNum = tdArr.length;
            }

            tdArr.forEach((otem,oindex,tdArr)=>{
                let rArr = /rowspan\s*=\s*\"(\d+)\"/.exec(otem);
                let cArr = /colspan\s*=\s*\"(\d+)\"/.exec(otem);
                let str = otem.replace(/<td[^>]*>/i,'').slice(0,-5);

                if(rowspanList[x+'-'+y]){
                    rowspanList[x+'-'+y] === 1 ? cellArr.push( table.createCell().addRowSpanOn() )
                        :cellArr.push( table.createCell().addRowSpanOn().addColSpan(rowspanList[x+'-'+y]) );
                }

                if(!(rArr || cArr)){
                    let cell = table.createCell();

                    str.split('<br />').forEach((ptem)=>{
                        cell.addParagraph(paragraph(ptem).setSpace({
                            'after':0
                        }))
                    });

                    cellArr.push( cell );
                    y++;
                }else if(rArr && !cArr){
                    cellArr.push( table.createCell().addParagraph(paragraph(str)).addRowSpanStart() );
                    for(let i=0; i<rArr[1]; i++){
                        rowspanList[x+'-'+(y+i)] = 1;
                    }
                    y++;
                }else if(!rArr && cArr ){
                    cellArr.push( table.createCell().addParagraph(paragraph(str)).addColSpan(cArr[1]) );
                    y = y + cArr[1];
                }else{
                    cellArr.push( table.createCell().addParagraph(paragraph(str)).addRowSpanStart().addColSpan(cArr[1]) );
                    for(let i=0; i<rArr[1]; i++){
                        rowspanList[x+'-'+(y+i)] = rArr[1];
                    }
                    y = y + cArr[1];
                }


            });
            trObj.addCell(cellArr);
            tb.addRow(trObj);
            x++;
        });



        return tb;
    };


    static checkType(valueName,value,type){
        if(typeof value !== type){
            throw new Error(`${valueName} type error`)
        }
    };


    static setSpace(p,styleObj){
        if(p['setSpace']){
            p.setSpace(styleObj)
        }

    }

    static convertXML2XMLObj(xml){
        let label,tag,arr;
        xml = xml.replace(/<([^>]+)>/,($$,$1)=>{
            label = $1.trim().split(' ')[0];
            return '';
        });

        tag = `</${label}>`;


        arr = xml.split(tag);
        arr.pop();

        return {
            [label]:{
                _cdata:arr.join(tag)
            }
        }

    }
}




module.exports = Docx;

