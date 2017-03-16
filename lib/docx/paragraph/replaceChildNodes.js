/*
*
*
* */

module.exports = function(parentNode=[],childNodeName='',childNode={}){
    let has = parentNode.some((item,index)=>{
        if(Object.keys(item)[0] === childNodeName){
            parentNode.splice(index,1,childNode);
            return true;
        }
    });
    if(!has){
        parentNode.push(childNode);
    }
};