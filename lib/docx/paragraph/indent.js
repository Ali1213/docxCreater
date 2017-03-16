module.exports = function (num) {
    'use strict';

    return {
        'w:ind': [{
            _attr: {
                'w:firstLineChars': 50*num,
                'w:firstLine': 110*num
            }
        }]
    };
};