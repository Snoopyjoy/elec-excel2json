const { app, BrowserWindow, ipcMain, Menu } = require('electron')
const { convert } = require('./xlsx2json');

const template = [
  {
    label: 'File',
    submenu: [
      { role: 'quit' }
    ]
  }
];

const menu = Menu.buildFromTemplate(template)
Menu.setApplicationMenu(menu);

function createWindow () {   
  // 创建浏览器窗口
  let win = new BrowserWindow({
    width: 800,
    height: 700,
    webPreferences: {
      nodeIntegration: true
    }
  })

  // 加载index.html文件
  win.loadFile('index.html')

  win.on('closed', () => {
    // 取消引用 window 对象，如果你的应用支持多窗口的话，
    // 通常会把多个 window 对象存放在一个数组里面，
    // 与此同时，你应该删除相应的元素。
    win = null
  })
}

app.on('window-all-closed', () => {
    // 在 macOS 上，除非用户用 Cmd + Q 确定地退出，
    // 否则绝大部分应用及其菜单栏会保持激活。
    if (process.platform !== 'darwin') {
      app.quit()
    }
  })
  
app.on('activate', () => {
  // 在macOS上，当单击dock图标并且没有其他窗口打开时，
  // 通常在应用程序中重新创建一个窗口。
  if (win === null) {
    createWindow()
  }
})

app.on('ready', createWindow)

ipcMain.on( 'exportTable', ( event, {
  side, pack, outputPath, tables
} )=>{
  convert( tables, outputPath, side, pack ).then( ()=>{
    event.reply('expSucc');
  } ).catch( (err)=>{
    event.reply('expFail', err.stack );
  } );
} )