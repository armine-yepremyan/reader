const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

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

const reader = (fileToRead) => {

   var workbook = XLSX.readFile(fileToRead);

   workbook.SheetNames.forEach(sheet => {
      let rowObject = XLSX.utils.sheet_to_json(
         workbook.Sheets[sheet], {skipHeader: true}
      )
      console.log("rowObject: ", rowObject.length);
   
      //TODO: complite column Names properly
      const column = {
		  A: "Path",
		  B: "Name",
		  C: "Age",
		  D: "File"
      };

      rowObject.forEach(obj => {
		  let targetDir = [];
		  let targetFile = '';
		  /* obj is the row object. example: 
		  {
		  	Path: '/home/armine/Desktop/anotherFolder/folderA/folderC', 
			Name: 'A.txt', 
			Age: 5}
		  */
		  for(let i in column) {
			  //checking if column is empty or not and as an example chosing filename
			  if(obj.hasOwnProperty(column[i]) && i !== 'D') {
				  targetDir.push(obj[column[i]]);
			  }
			  
			  if(obj.hasOwnProperty(column[i]) && i === 'D') {
				  targetFile = obj[column[i]];
			  }
		  }
		  
		  //TODO: change _root path
		  const _root = '/home/armine/Desktop/';
		  let _path = _root + targetDir.join('/');
		  createPath(_path);
		  createFile(_path, targetFile, '.txt');
      })

   });
}

const argv = process.argv;
console.log(argv[2]);

reader(argv[2]);