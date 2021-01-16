	const XLSX = require('xlsx');
	//const fs = require('fs');
	const fs = require(	'fs-extra');
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



const reader = (fileToRead, _root, localNASpath) => {

	    var workbook = XLSX.readFile(fileToRead);

	    workbook.SheetNames.forEach(sheet => {
	    let rowObject = XLSX.utils.sheet_to_json(
	       workbook.Sheets[sheet], {skipHeader: true}
	    )
		  let	rowN = rowObject.length
	    console.log("number of rows to analyze: ", rowObject.length);

			console.log(`indirizzo NAS ${localNASpath}`)


			/* Get FIRST worksheet */
			var worksheet = workbook.Sheets[workbook.SheetNames[0]];
			/* For each row */
			for (var r = 1; r <= rowN; r++) {

			const monthNames = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno","Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
			const monthNumbers = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];
			const dayNumbers = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"];

			console.log(`analyzing row number ${r}`)
			/* Get the desired values */
			 var exDate = (worksheet[`B${r+1}`] ? worksheet[`B${r+1}`].v : undefined);
			 //extract yyyy, mm, dd from date read in xls cell
			 const { getJsDateFromExcel } = require("excel-date-to-js"); // USED package npm i excel-date-to-js
			 var exDateConv = getJsDateFromExcel(exDate);
			 var exYear = exDateConv.getFullYear();
			 var exMonth = monthNumbers[exDateConv.getMonth()];  var exMonthName = monthNames[exDateConv.getMonth()];
			 //var exDay = exDateConv.getDate();
			 var exDay = dayNumbers[exDateConv.getDate()-1];
			 var pLastName = (worksheet[`F${r+1}`] ? worksheet[`F${r+1}`].v : undefined);
			 var pName = (worksheet[`G${r+1}`] ? worksheet[`G${r+1}`].v : undefined);


			 var typeDep = (worksheet[`J${r+1}`] ? worksheet[`J${r+1}`].v : undefined);
			 if (typeDep != undefined) {
				 var pathSubFolderOut = '';
        if (typeDep.includes("Ambulatoriale")){
					//pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`; //ex: 2021_01_04_PostProtesi_
					typeDep = 1; //console.log(`type patient ${typeP}`)
				}
				else if (typeDep.includes("Nutrizionale")){
					//pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`; //ex: 2021_01_04_PostProtesi_
					typeDep = 2; //console.log(`type patient ${typeP}`)
				}
				else if (typeDep.includes("RRF")){
					pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`; //ex: 2021_01_04_PostProtesi_
					typeDep = 3; //console.log(`type patient ${typeP}`)
				}
				else if (typeDep.includes("CMF")){
						pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`; //ex: 2021_01_04_PostProtesi_
						typeDep = 4; //console.log(`type patient ${typeP}`)
					}
					else {
					console.log(`error type patient`)
					}
				}


			 var typeP = (worksheet[`L${r+1}`] ? worksheet[`L${r+1}`].v : undefined);
			 if (typeP != undefined) {
				 var pathSubFolderOut = '';
				 if (typeP.includes("Pre Protesi")){
					 pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`;
					 typeP = 1; //console.log(`type patient ${typeP}`)
				} else if (typeP.includes("Post Protesi")){
					pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`;
					typeP = 2; //console.log(`type patient ${typeP}`)
				} else if (typeP.includes("Protesi")){
					pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`;
					typeP = 3; //console.log(`type patient ${typeP}`)
				}
				else if (typeP.includes("Nutrizionale I")){
					//pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`
					typeP = 4; //console.log(`type patient ${typeP}`)
				}
				else if (typeP.includes("Nutrizionale D")){
					//pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`;
					typeP = 5; //console.log(`type patient ${typeP}`)
				}
					else if (typeP.includes("Nutrizionale")){
						//pathSubFolderOut=`${exYear}_${exMonth}_${exDay}_${typeP}_`;
						typeP = 5; //console.log(`type patient ${typeP}`)
					}
					else {
					console.log(`error type patient`)
					}
				}




				var typeProsthesis = (worksheet[`M${r+1}`] ? worksheet[`M${r+1}`].v : undefined);

				// if typeProsthesis is "Ginocchio" then jointProsthesis=1; if is "Anca" then jointProsthesis=2 ; else = undefined
				 if (typeProsthesis != null) {
					 if (typeProsthesis.includes("Ginocchio")){
						 pathSubFolderOut=`${pathSubFolderOut}Ginocchio`; //ex: 2021_01_04_PreProtesi_Ginocchio
						 var jointProsthesis = 1; //console.log(`type prothesis ${jointProsthesis}`)
					} else if (typeProsthesis.includes("Anca")){
						pathSubFolderOut=`${pathSubFolderOut}Anca`; //ex: 2021_01_04_PreProtesi_Anca
						var jointProsthesis = 2; //console.log(`type prothesis ${jointProsthesis}`)
					} else {
						pathSubFolderOut=`${pathSubFolderOut}XXXX`; //
						console.log(`error joint prothesis`)
						}
					}

        // if typeProsthesis is "DX" then sideProsthesis=1; if is "SX" then sideProsthesis=2 ; ; if is "DX e SX" then sideProsthesis=3; else = undefined
				if (typeProsthesis != null) {
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
					console.log(`error side prothesis`)
					}
				}

				// if ex is a Post, but not Pre present (for future: check on the colomnus of xls controls and/or check on NAS)
       	 if ((typeP == 1 || typeP == 2) && typeProsthesis == null) {
						 pathSubFolderOut=`${pathSubFolderOut}XXXX`; //ex:
						 console.log(`not Pre present`)
					 }



				 var exA = (worksheet[`P${r+1}`] ? worksheet[`P${r+1}`].v : undefined);
				 var exAp = (worksheet[`Q${r+1}`] ? worksheet[`Q${r+1}`].v : undefined);
				 var exR = (worksheet[`R${r+1}`] ? worksheet[`R${r+1}`].v : undefined);
				 var exCMJS = (worksheet[`T${r+1}`] ? worksheet[`T${r+1}`].v : undefined);
				 var exF = (worksheet[`U${r+1}`] ? worksheet[`U${r+1}`].v : undefined);
				 var exP = (worksheet[`W${r+1}`] ? worksheet[`W${r+1}`].v : undefined);
				 var exE = (worksheet[`X${r+1}`] ? worksheet[`X${r+1}`].v : undefined);
				 var exD = (worksheet[`Y${r+1}`] ? worksheet[`Y${r+1}`].v : undefined);
				 var exC = (worksheet[`Z${r+1}`] ? worksheet[`Z${r+1}`].v : undefined);
				 var exTUG = (worksheet[`AA${r+1}`] ? worksheet[`AA${r+1}`].v : undefined);
				 var ex6MWT = (worksheet[`AB${r+1}`] ? worksheet[`AB${r+1}`].v : undefined);
	       var exPreYN = (worksheet[`AK${r+1}`] ? worksheet[`AK${r+1}`].v : undefined);
				 // if EX is Post and Exists Pre, then exPreYN==1 ; else  exPreYN==0 --
				 // add control: type Pre == type post
				 exPreYN = (exPreYN >0 ? 1 : 0);

				 // array containing a number of elements equal to the number of file to createFile
				 // each element contains the part of the filename to add
				 exNamesToCreate = [];
				 if (exA != null) {
					 exNamesToCreate.push('AnalisiCammino_CinDin');
				 }
				 if (exAp != null) {
					 exNamesToCreate.push('AnalisiCammino_CinDinEMG');
					 exNamesToCreate.push('AnalisiCammino_EMG');
				 }
				 if (exR != null) {
					 exNamesToCreate.push('AnalisiCammino_CinDin');
					 exNamesToCreate.push('AnalisiCamminoTreadmill_XXkmh_CinDin');
					 exNamesToCreate.push('AnalisiCorsaTreadmill_XXkmh_CinDin');
				 }
				 if (exCMJS != null) {
					 exNamesToCreate.push('AnalisiSquat_Bipodalico_CinDin');
				 }
				 if (exF != null) {
					 exNamesToCreate.push('AnalisiCervicale_Cin');
				 }
				 if (exP != null) {
					 exNamesToCreate.push('AnalisiOrtostasi&Baropodometria');
				 }
				 if (exE != null) {
					 exNamesToCreate.push('Baropodometria');
					 exNamesToCreate.push('Walk');
				 }
				 if (exD != null) {
					 exNamesToCreate.push('Baropodometria');
				 }
				 if (exC != null) {

					 switch (typeP) {
						 case 1:
						        exNamesToCreate.push('Walk_PreProtesi');
						        break;
						case 2:
						        exNamesToCreate.push('Walk_PostProtesi');
					          exNamesToCreate.push('Walk_Protesi_ConfrontoPreVsPost'); //add control: create only if Pre present
										break;
						case 4:
										exNamesToCreate.push('Walk&TUG_Ingresso');
										break;
						case 5:
										exNamesToCreate.push('Walk&TUG_Dimissione');
										break;
						case undefined:
						        exNamesToCreate.push('Walk');;
										break;
					  }
				}


				 if (exTUG != null && (typeP == undefined || [1,2,4,5].indexOf(typeP) == -1)) {
					 exNamesToCreate.push('TUG');
				 }
				 if (ex6MWT != null) {
					 exNamesToCreate.push('6MWT');
				 }



				 	//create Dir and Files
					let pathFileOut = '';
					let pathFolderOut = '';
					let fileNameOut = '';
					pathFileOut =`${_root}${exDay} ${exMonthName} ${exYear}`;
					pathFolderOut =`${pathFileOut}\\${pLastName.toUpperCase()} ${pName.toUpperCase()}\\`;
					fileNameOut=`${exYear}_${exMonth}_${exDay}_${pLastName.replace(/\s/g, '')}_${pName.replace(/\s/g, '')}`;




				//  if in NAS folder "pLastName pName" in in localNASpath
				console.log(`searching path in NAS: ${`${localNASpath}\\${pLastName.toUpperCase()} ${pName.toUpperCase()}\\`}`);
				// then copy folder from NAS to pathFileOut
				// else createPath(pathFolderOut)
				if (fs.existsSync(`${localNASpath}\\${pLastName.toUpperCase()} ${pName.toUpperCase()}\\`)) {
					console.log('exists in NAS');
					fs.copySync(`${localNASpath}\\${pLastName.toUpperCase()} ${pName.toUpperCase()}\\`, pathFolderOut)
				}
				else {
					console.log('NOT exists in NAS');
					createPath(pathFolderOut)
				}



	      //  console.log(`nome del path ${pathFileOut}`);				//	console.log(`nome del path ${pathFolderOut}`);				//	console.log(`nome del file ${fileNameOut}`);
					if (typeP != undefined && typeof pathSubFolderOut !== 'undefined') {
						createPath(`${pathFolderOut}\\${pathSubFolderOut}`)
				  }

        //  for each date created, creat a "base empty folder/file combination" for Pre and Post ex
        createPath(`${pathFileOut}\\xxx`)
        createPath(`${pathFileOut}\\${exYear}_${exMonth}_${exDay}_PreProtesi_`)
        createPath(`${pathFileOut}\\${exYear}_${exMonth}_${exDay}_PostProtesi_`)
				createFile(pathFileOut, `_AnalisiCammino_CinDin_${exYear}_${exMonth}_${exDay}`,`.txt`);
				createFile(pathFileOut, `_Walk_PostProtesi_${exYear}_${exMonth}_${exDay}`,`.txt`);
				createFile(pathFileOut, `_Walk_PreProtesi_${exYear}_${exMonth}_${exDay}`,`.txt`);
				createFile(pathFileOut, `_Walk_Protesi_ConfrontoPreVsPost_${exYear}_${exMonth}_${exDay}`,`.txt`);

				for (let i = 0; i < exNamesToCreate.length; i++) {
				  	createFile(pathFileOut, `${fileNameOut}_${exNamesToCreate[i]}`,`.txt`);
					}


	  }
	});
	}



	const argv = process.argv;
	console.log(argv[3]);

	reader(argv[2], argv[3], argv[4]);
