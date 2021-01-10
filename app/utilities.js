const path = require('path');
const fs = require('fs');


//Constants
const monthNames = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
// const monthNumbers = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];
// const dayNumbers = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"];
const OS = {
    LINUX: 'linux',
    WINDOWS: 'win32',
};



// halper functions
const createPath = (targetDir) => {
    const sep = path.sep;
    const initDir = path.isAbsolute(targetDir) ? sep : '';
    targetDir.split(sep).reduce((parentDir, childDir) => {
        const curDir = path.resolve(parentDir, childDir);
        if (!fs.existsSync(curDir)) {
            fs.mkdirSync(curDir);
        }
        return curDir;
    }, initDir);
}

const createFile = (targetDir, fileName, extension) => {
    console.log("createFile: ", `${targetDir}/${fileName}${extension}`);
    fs.closeSync(fs.openSync(`${targetDir}/${fileName}${extension}`, 'w'));
}


module.exports.monthNames = monthNames;
module.exports.OS = OS;
// module.exports.monthNumbers = monthNumbers;
// module.exports.dayNumbers = dayNumbers;

module.exports.createPath = createPath;
module.exports.createFile = createFile;
 