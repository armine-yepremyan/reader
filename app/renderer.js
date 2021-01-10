const ipc = require('electron').ipcRenderer;
const fs = require('fs');

let _inputFilePath = '';
let _outputFolderPath = '';

document.getElementById('fileToRead').addEventListener('click', (event) => {
    ipc.send('open-file-dialog-for-file', 'ping');
});

document.getElementById('selectedFolder').addEventListener('click', (event) => {
    ipc.send('open-file-dialog-for-folder', 'pong');
});

document.querySelector('form').addEventListener('submit', e => {
    ipc.send('create', { _inputFilePath, _outputFolderPath });
});

ipc.on('selected-file', (event, inputFilePath) => {
    _inputFilePath = inputFilePath.filePaths[0];
    document.getElementById("filePath").placeholder = _inputFilePath;
});

ipc.on('selected-folder', (event, outputFolderPath) => {
    _outputFolderPath = outputFolderPath.filePaths[0];
    document.getElementById("folderPath").placeholder = _outputFolderPath;
});