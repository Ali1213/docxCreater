


# how to start

## simple example

there is a simple example:

```javascript
var d = new Docx();
d.addParagraph('there is content');
d.render(/*filepath*/);
```

also, it equals to the next:
```javascript
var d = new Docx();
d.addParagraph('there is content')
    .render(/*filepath*/);
```

## something you must know

In docx file, the fundamental elements is paragraph and table.
you can use `d.addParagraph()`to add a paragraph.
you can use `d.addTable()`to add a table.
If you want to add a Image or text, you must  addParagraph at first.

## d.prototype.setPages()

If you want to set pages,you must you this;
```javascript
d.setPage({
        type = "", //string 'A3','A4'
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
    })
```

## d.prototype.addParagraph(string,opt);
```
opt
    /*
     * indent:段首缩进字符，默认是0个字符
     * alignment:段落对齐方式，"right","left","center","justified",4选1，默认left
     * */
```

## d.prototype.addText(string,opt);
```
addText(str,{
        fontSize = 12,
        subScript = false,
        superScript = false,
        bold = false,
        italic = false,
        underline = false,
        strike = false,
    }={})
    //fontSize : number/string 字体大小
    //subScript : boolean 是否是下标
    //superScript : boolean 是否是下标
    //bold : boolean 是否是黑体
    //italic : boolean 是否是下标
    //underline : boolean 是否是下标
    //strike : boolean 是否是下标
```


## d.prototype.addLink(string,link);

## d.prototype.addImage(filepath,width,height,type = "inset");
```
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
```


## d.prototype.addTable(options);
```
options.HTMLTable //string HTML formating table string
```


## d.prototype.render(options,cb);
docx generate.
```
options:
    creator
    description
    title
    subject
    keywords
    lastModifiedBy
    revision
cb: function
 two arguments,first is error,second is result;
```


# 更新日志

+ 2017/3/28

    - 修复创建段落时入不放入文字会造成的bug

+ 2017/3/22

    - 修復addText時設置字體的問題
    - 修复若干小问题


+ 2017/3/13

    - 更新一个可以直接插入合法的xml接口。(!慎用)
    - 支持设置段前段后间距，以及行间距。

+ 2017/3/10
    - 修正HTML形式转table的bug
    - 临时增加一段直接插入document.xml的obj接口,方便将试卷密封圈的文本插进去
    - 增加设置页面的 上下左右的边距
    - 增加table的边框属性,暂未支持设置接口


+ 2017/3/9
    - 增加HTML形式的table转成table
    - 增加对字体上下标和黑体，斜体，带下划线，带方框的支持
    - 增加简单分栏，增加页面设置（A4，A3）
+ 2017/3/7
    - 增加对段落字体大小的统一设置
+ 2017/3/1
    - 支持JPG,JPEG图片
    - 支持PNG图片
    - 支持gif图片
    - 支持段落的对齐方式 "right","left","center","justified";
    - 图片的几种方式:嵌入:segments;四周环绕:square;紧密环绕:tight;穿越型环绕:through;上下型环绕:topAndBottom;衬于文字下方:underText;衬于文字上方:aboveText;然而目前由于无法定位图片和文字的位置，所以这些功能暂时无法完善
+ 2017/2/28
    - 增加段落方法createParagraph
    - 对已增加的段落增加文字addText
    - 增加对超链接文字的支持




