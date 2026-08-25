'use strict';

const { app, BrowserWindow, Menu, session, shell } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const { isInternalNav, isOpenableExternal } = require('./url-policy');

const START_URL = 'https://outlook.office.com/mail';

// --- Link routing: single choke point for every webContents the app creates ---

// Strip the Electron marker so Microsoft's browser gating serves the full page.
// Only webContents.setUserAgent changes navigator.userAgent (Electron 44 ignores
// the --user-agent Chromium switch); session.setUserAgent covers request headers.

function openInBrowser(url) {
  if (!isOpenableExternal(url)) {
    console.error(`[url-policy] dropped non-openable URL: ${url}`);
    return;
  }
  shell.openExternal(url);
}

app.on('web-contents-created', (_e, wc) => {
  wc.setWindowOpenHandler(({ url }) => {
    openInBrowser(url);
    return { action: 'deny' };
  });
  wc.on('will-navigate', (e, url) => {
    if (!isInternalNav(url)) {
      e.preventDefault();
      openInBrowser(url);
    }
  });
});

// --- Window state persistence (~zero deps) ---

function stateFile() {
  return path.join(app.getPath('userData'), 'window-state.json');
}

const DEFAULTS = { width: 1280, height: 800 };

function loadWindowState() {
  try {
    const s = JSON.parse(fs.readFileSync(stateFile(), 'utf8'));
    if (typeof s.width !== 'number' || typeof s.height !== 'number') return DEFAULTS;
    return s;
  } catch {
    return DEFAULTS;
  }
}

function saveWindowState(win) {
  const s = win.isMaximized()
    ? { ...loadWindowState(), isMaximized: true }
    : { ...win.getBounds(), isMaximized: false };
  try {
    fs.writeFileSync(stateFile(), JSON.stringify(s));
  } catch (err) {
    console.error('[window-state] save failed:', err.message);
  }
}

// --- Main window & lifecycle ---

let mainWindow = null;

function createWindow(strippedUA) {
  const state = loadWindowState();
  mainWindow = new BrowserWindow({
    x: state.x,
    y: state.y,
    width: state.width,
    height: state.height,
    minWidth: 480,
    minHeight: 320,
    title: 'Outlook',
    backgroundColor: '#f3f2f1',
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      partition: 'persist:outlook',
      sandbox: true,
      contextIsolation: true,
    },
  });
  if (state.isMaximized) mainWindow.maximize();
  mainWindow.once('ready-to-show', () => mainWindow.show());
  mainWindow.on('close', () => saveWindowState(mainWindow));
  mainWindow.on('closed', () => { mainWindow = null; });
  mainWindow.webContents.setUserAgent(strippedUA);
  mainWindow.loadURL(START_URL);
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (!mainWindow) return;
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  });

  app.whenReady().then(() => {
    const strippedUA = app.userAgentFallback
      .replace(/ Electron\/\S+/, '')
      .replace(/ {2,}/g, ' ');

    const sess = session.fromPartition('persist:outlook');
    sess.setUserAgent(strippedUA);
    sess.setPermissionRequestHandler((_wc, permission, cb) => {
      cb(['notifications', 'media', 'fullscreen', 'clipboard-sanitized-write'].includes(permission));
    });

    Menu.setApplicationMenu(Menu.buildFromTemplate([
      {
        label: 'App',
        submenu: [{ role: 'quit' }],
      },
      {
        label: 'Edit',
        submenu: [
          { role: 'undo' }, { role: 'redo' }, { type: 'separator' },
          { role: 'cut' }, { role: 'copy' }, { role: 'paste' },
          { role: 'selectAll' },
        ],
      },
      {
        label: 'View',
        submenu: [
          { role: 'reload' }, { role: 'forceReload' }, { role: 'toggleDevTools' },
          { type: 'separator' },
          { role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' },
        ],
      },
    ]));

    createWindow(strippedUA);
  });

}
