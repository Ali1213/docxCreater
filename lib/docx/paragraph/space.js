




module.exports = function (opt){
    let attr = {};

    if(!isNaN(opt.after)){
        attr['w:afterLines'] = opt.after;
        attr['w:after'] = parseInt(opt.after * 3.12);
    }

    if(!isNaN(opt.before)){
        attr['w:beforeLines'] = opt.before;
        attr['w:before'] = parseInt(opt.before * 3.12);
    }

    // attr['w:lineRule'] = 'auto';

    if(!isNaN(opt.line)){
        attr['w:line'] = opt.line;
        attr['w:lineRule'] = opt.lineType ? opt.lineType : 'auto' ;
    }


    return {
            'w:spacing':[{
                _attr:attr
            }]
    }
}