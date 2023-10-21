const xlsx = require('xlsx');
var fs = require('fs');
const path = require('path');

function xtj(file){
    const fil = xlsx.readFile(file.path);
    const sheetNames = fil.SheetNames;
    const totalSheets = sheetNames.length;
    let parsedData = [];
    for (let i = 0; i < totalSheets; i++) {
        const tempData = xlsx.utils.sheet_to_json(fil.Sheets[sheetNames[i]],{ header: 1 });
        nam = tempData[0];
        tempData.shift();
        parsedData.push({header:nam,
                        body:tempData});
    }
    fn = "/uploads/json/"+file.originalname.substring(0,file.originalname.length-4)+"json";
    fs.writeFile(path.join(__dirname+"/../", fn), JSON.stringify(parsedData), function(err) {
    if (err) {
        console.log(err);
    }
    });
}

module.exports = {xtj};