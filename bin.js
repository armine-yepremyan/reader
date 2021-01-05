const read = require('.');

const argv = process.argv;
let fileToRead = argv[2];
//let pathToSave = argv[3];
console.log(argv);
read(fileToRead);
