const XLSX = require('xlsx');
const { getJsDateFromExcel } = require("excel-date-to-js");
const fs = require('fs');
const path = require('path');

//Constants
const monthNames = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];

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
    fs.closeSync(fs.openSync(`${targetDir}/${fileName}${extension}`, 'w'));
}

// main function
const createFilesDirectories = (inputFilePath, outputFolderPath, callback) => {
    let err = null;
    try {
        console.info("OS is ", process.platform);
        const slash = process.platform === OS.WINDOWS ? "\\" : '/';
        if(outputFolderPath.charAt(outputFolderPath.length - 1) !== slash) outputFolderPath += slash;
        var workbook = XLSX.readFile(inputFilePath);
        
        /* Get FIRST worksheet */
        var worksheet = workbook.Sheets[workbook.SheetNames[0]];
        let rowObject = XLSX.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[0]], {skipHeader: true}
        );
        let	rowN = rowObject.length;
        console.info("number of rows to analyze: ", rowObject.length);
        
        /* For each row */
        for(let r = 1; r <= rowN; r++) {
            console.info(`analyzing row number ${r}`);
            /* Get the desired values */
            let exDate = (worksheet[`B${r+1}`] ? worksheet[`B${r+1}`].v : undefined);
            //extract yyyy, mm, dd from date read in xls cell
            // USED package npm i excel-date-to-js
            let exDateConv = getJsDateFromExcel(exDate);
            
            let exYear = exDateConv.getFullYear();
            
            let exMonthName = monthNames[exDateConv.getMonth()];
            
            let _m = exDateConv.getMonth() + 1;
            let exMonth = _m < 10 ? `0${_m.toString()}` : _m.toString();
            
            let _d = exDateConv.getDate();
            let exDay = _d < 10 ? `0${_d.toString()}` : _d.toString();
            
            let pLastName = (worksheet[`F${r+1}`] ? worksheet[`F${r+1}`].v : undefined);
            let pName = (worksheet[`G${r+1}`] ? worksheet[`G${r+1}`].v : undefined);
            let typeP = (worksheet[`L${r+1}`] ? worksheet[`L${r+1}`].v : undefined);
            
            // if typeP is "preProtesi" then typeP=1; if is "PostProtesi" then typeP=2 ; else = undefined
            if (typeP !== undefined) {
                var pathSubFolderOut = '';
                if (typeP.includes("Pre Protesi")){
                    pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`; //ex: 2021_01_04_PreProtesi_
                    typeP = 1; //console.log(`type patient ${typeP}`)
                } else if (typeP.includes("Post Protesi")){
                    pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`; //ex: 2021_01_04_PostProtesi_
                    typeP = 2; //console.log(`type patient ${typeP}`)
                } else {
                    console.log(`error type patient`); //TODO: on message box?
                }
            }
            
            var typeProsthesis = (worksheet[`M${r+1}`] ? worksheet[`M${r+1}`].v : undefined);
        
            // if typeProsthesis is "Ginocchio" then jointProsthesis=1; if is "Anca" then jointProsthesis=2 ; else = undefined
            if (typeProsthesis !== undefined) {
                if (typeProsthesis.includes("Ginocchio")){
                    pathSubFolderOut=`${pathSubFolderOut}Ginocchio`; //ex: 2021_01_04_PreProtesi_Ginocchio
                    var jointProsthesis = 1; //console.log(`type prothesis ${jointProsthesis}`)
                } else if (typeProsthesis.includes("Anca")){
                    pathSubFolderOut=`${pathSubFolderOut}Anca`; //ex: 2021_01_04_PreProtesi_Anca
                    var jointProsthesis = 2; //console.log(`type prothesis ${jointProsthesis}`)
                } else {
                    pathSubFolderOut=`${pathSubFolderOut}XXXX`; //
                    console.log(`error joint prothesis`); //TODO: error!messagebox?
                }
            }
        
            // if typeProsthesis is "DX" then sideProsthesis=1; if is "SX" then sideProsthesis=2 ; ; if is "DX e SX" then sideProsthesis=3; else = undefined
            if (typeProsthesis !== undefined) {
                if (typeProsthesis.includes("DX")){
                    pathSubFolderOut=`${pathSubFolderOut}DX`; //ex: 2021_01_04_PreProtesi_GinocchioDX
                    var sideProsthesis = 1; //console.log(`side prothesis ${sideProsthesis}`)
                } else if (typeProsthesis.includes("SX")){
                    pathSubFolderOut=`${pathSubFolderOut}SX`; //ex: 2021_01_04_PreProtesi_GinocchioDX
                    var sideProsthesis = 2; //console.log(`side prothesis ${sideProsthesis}`)
                } else if (typeProsthesis.includes("DX e SX")){
                    pathSubFolderOut=`${pathSubFolderOut}DXeSX`; //ex: 2021_01_04_PreProtesi_GinocchioDX
                    var sideProsthesis = 3; //console.log(`side prothesis ${sideProsthesis}`)
                } else {
                    pathSubFolderOut=`${pathSubFolderOut}XXXX`; //
                    console.log(`error side prothesis`); //TODO: error!messagebox?
                }
            }
        
            // if ex is a Post, but not Pre present (for future: check on the colomnus of xls controls and/or check on NAS)
            if ((typeP == 1 || typeP == 2) && typeProsthesis == null) {
                pathSubFolderOut=`${pathSubFolderOut}XXXX`; //ex:
                
                console.log(`not Pre present`); //TODO: error!messagebox?
            }
        
            let exA = (worksheet[`P${r+1}`] ? worksheet[`P${r+1}`].v : undefined);
            let exAp = (worksheet[`Q${r+1}`] ? worksheet[`Q${r+1}`].v : undefined);
            let exR = (worksheet[`R${r+1}`] ? worksheet[`R${r+1}`].v : undefined);
            let exCMJS = (worksheet[`T${r+1}`] ? worksheet[`T${r+1}`].v : undefined);
            let exF = (worksheet[`U${r+1}`] ? worksheet[`U${r+1}`].v : undefined);
            let exP = (worksheet[`W${r+1}`] ? worksheet[`W${r+1}`].v : undefined);
            let exE = (worksheet[`X${r+1}`] ? worksheet[`X${r+1}`].v : undefined);
            let exD = (worksheet[`Y${r+1}`] ? worksheet[`Y${r+1}`].v : undefined);
            let exC = (worksheet[`Z${r+1}`] ? worksheet[`Z${r+1}`].v : undefined);
            let exTUG = (worksheet[`AA${r+1}`] ? worksheet[`AA${r+1}`].v : undefined);
            let ex6MWT = (worksheet[`AB${r+1}`] ? worksheet[`AB${r+1}`].v : undefined);
            let exPreYN = (worksheet[`AK${r+1}`] ? worksheet[`AK${r+1}`].v : undefined);
            // if EX is Post and Exists Pre, then exPreYN==1 ; else  exPreYN==0 --
            // add control: type Pre == type post
            exPreYN = (exPreYN > 0 ? 1 : 0);
        
            // array containing a number of elements equal to the number of file to createFile
            // each element contains the part of the filename to add
            let exNamesToCreate = [];
            if (exA !== undefined) {
                exNamesToCreate.push('AnalisiCammino_CinDin');
            }
            if (exAp !== undefined) {
                exNamesToCreate.push('AnalisiCammino_CinDinEMG');
                exNamesToCreate.push('AnalisiCammino_EMG');
            }
            if (exR !== undefined) {
                exNamesToCreate.push('AnalisiCammino_CinDin');
                exNamesToCreate.push('AnalisiCamminoTreadmill_CinDin');
                exNamesToCreate.push('AnalisiCorsaTreadmill_CinDin');
            }
            if (exCMJS !== undefined) {
                exNamesToCreate.push('AnalisiSquatBipodalico_CinDin');
            }
            if (exF !== undefined) {
                exNamesToCreate.push('AnalisiCervicale_Cin');
            }
            if (exP !== undefined) {
                exNamesToCreate.push('AnalisiOrtostasi&Baropodometria');
            }
            if (exE !== undefined) {
                exNamesToCreate.push('Baropodometria');
                exNamesToCreate.push('TUG');
            }
            if (exD !== undefined) {
                exNamesToCreate.push('Baropodometria');
            }
            if (exTUG !== undefined) {
                exNamesToCreate.push('TUG');
            }
            if (ex6MWT !== undefined) {
                exNamesToCreate.push('6MWT');
            }
            
        
            //create Dir and Files
            let pathFileOut = '';
            let pathFolderOut = '';
            let fileNameOut = '';
            pathFileOut =`${outputFolderPath}${exDay} ${exMonthName} ${exYear}`;
            pathFolderOut =`${pathFileOut}${slash}${pLastName.toUpperCase()} ${pLastName.toUpperCase()}${slash}`;
            fileNameOut=`${exYear}_${exMonth}_${exDay}_${pLastName}_${pName}`;
            
            createPath(pathFolderOut);
            if (typeof pathSubFolderOut !== 'undefined') {
                createPath(`${pathFolderOut}${slash}${pathSubFolderOut}`)
            }
        
            exNamesToCreate.forEach(_exNamesToCreate => {
                createFile(pathFileOut, `${fileNameOut}_${_exNamesToCreate}`,`.txt`);
            })
            // for (let i = 0; i < exNamesToCreate.length; i++) {
            //     console.log("create1: ", `${fileNameOut}_${exNamesToCreate[i]}`);
            //     createFile(pathFileOut, `${fileNameOut}_${exNamesToCreate[i]}`,`.txt`);
            // }
        }
        
    } catch (e) {
        err = {
            message: "Something went wrong!", //maybe loaded wrong file
            detail: `${e}`
        }
    }
    return callback(err, "Created");        
        
}

module.exports = createFilesDirectories;
 