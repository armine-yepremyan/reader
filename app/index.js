const { app, dialog, BrowserWindow } = require('electron')
const ipc = require('electron').ipcMain;
// const os = require('os');
const createFilesDirectories = require('./halper');


let mainWindow = null;

ipc.on('close-main-window', () => app.quit());

function createWindow () {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true
    }
  })

  mainWindow.loadFile('index.html')
  
  mainWindow.on('closed', () => mainWindow = null);
  
  ipc.on('open-file-dialog-for-folder', (event) => {
      dialog.showOpenDialog(mainWindow, {
          properties: ['openDirectory']
      }).then(files => {
          if (files) event.reply('selected-folder', files);
      });      
  })
  
  ipc.on('open-file-dialog-for-file', (event) => {
      dialog.showOpenDialog(mainWindow, {
          filters: [{ name: 'SpreedSheet files', extensions: ['xls', 'xlsx'] }],
          properties: ['openFile']
      }).then(files => {
          if (files) event.reply('selected-file', files);
      });      
  })
  
  ipc.on('create', (event, data) => {
      if(!data._inputFilePath || !data._outputFolderPath) {
          const options = {
              type: 'error',
              buttons: ['Ok'],
              defaultId: 0,
              title: 'Error!',
              message: 'Have you selected spreedsheet file and output folder?',
              detail: 'Please Select Input File and Output Folder.'
          };
          dialog.showMessageBox(mainWindow, options);
          return;
      }
      createFilesDirectories(data._inputFilePath, data._outputFolderPath, (err, data) => {
          let options = null;
          if(err) {
              options = {
                  type: 'error',
                  buttons: ['Ok'],
                  defaultId: 0,
                  title: 'Error!',
                  message: err.message,
                  detail: err.detail
              };
              
          } else {
              options = {
                  type: 'info',
                  buttons: ['Ok'],
                  defaultId: 0,
                  title: 'Done!',
                  message: 'All Files and Folders are created'
              };
          }
          dialog.showMessageBox(mainWindow, options);          
      });
  })
  
  
  
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
})
