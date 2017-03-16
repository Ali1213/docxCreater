const reNode = require("./replaceChildNodes");


const szCs = function(num){
    return {
        'w:szCs': [{
            _attr: {
                'w:val': num ? num*2:32
            }
        }]
    };
};


const sz = function(num){
    return {
        'w:sz': [{
            _attr: {
                'w:val': num ? num*2:32
            }
        }]
    };
};



module.exports = function(){
  return {
      'w:rPr':[{
          _attr:{}
      }],
      'fontSize':function(num){
          reNode(this['w:rPr'],'w:sz',sz(num));
          reNode(this['w:rPr'],'w:szCs',sz(num));
          return this;
      }
  }
};