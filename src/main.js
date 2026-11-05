'use strict';

const { app, BrowserWindow, Menu, session, shell } = require('electron');
const fs = require('node:fs');
const path = require('node:path');

const START_URL = 'https://outlook.office.com/mail';
const APP_ICON = path.join(__dirname, '..', 'assets', 'icons', 'outlook-for-linux-256.png');

// --- Link routing: single choke point for every webContents the app creates ---

const { isInternalNav, isOpenableExternal, isAuthHost } = require('./url-policy');

// Strip the Electron marker so Microsoft's browser gating serves the full page.
// Only webContents.setUserAgent changes navigator.userAgent (Electron 44 ignores
// the --user-agent Chromium switch); session.setUserAgent covers request headers.

function openInBrowser(url) {
  if (!isOpenableExternal(url)) {
    console.error(`[url-policy] dropped non-openable URL: ${url}`);
    return;
  }
  console.error(`[url-policy] routing to system browser: ${url}`);
  shell.openExternal(url);
}

app.on('web-contents-created', (_e, wc) => {
  wc.setWindowOpenHandler(({ url }) => {
    // Inside an auth flow the IdP may open popups (consent, FIDO, signup);
    // those must render in-app or the flow breaks. Everywhere else, route out.
    if (isAuthHost(wc.getURL()) && isOpenableExternal(url)) {
      return { action: 'allow' };
    }
    openInBrowser(url);
    return { action: 'deny' };
  });
  wc.on('will-navigate', (e, url) => {
    // While the current page is an auth host, the flow may bounce across
    // unpredictable IdP hosts (signup, per-tenant endpoints) — let https
    // navigations render in-app. Reading mail: everything non-internal
    // goes to the system browser.
    if (!isInternalNav(url) && !isAuthHost(wc.getURL())) {
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

if (process.defaultApp) {
  if (process.argv.length >= 2) {
    app.setAsDefaultProtocolClient("mailto", process.execPath, [
      path.resolve(process.argv[1]),
    ]);
  }
} else {
  app.setAsDefaultProtocolClient("mailto");
}


let mainWindow = null;
let composeWindow = null;
let composeUrl = null;

function createComposeWindow(url) {
  composeUrl = encodeURI("https://outlook.cloud.microsoft/mail/deeplink/compose?"+url.replace(/mailto:\/*/, "to="));
  composeWindow = new BrowserWindow({
    minWidth: 480,
    minHeight: 320,
    title: 'Outlook - Compose',
    icon: APP_ICON,
    backgroundColor: '#f3f2f1',
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      partition: 'persist:outlook',
      sandbox: true,
      contextIsolation: false,
    },
  });
  composeWindow.once('ready-to-show', () => composeWindow.show());
  composeWindow.on('closed', () => { composeWindow = null; });
  composeWindow.loadURL(composeUrl);
}

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
    icon: APP_ICON,
    backgroundColor: '#f3f2f1',
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      partition: 'persist:outlook',
      sandbox: true,
      contextIsolation: false,
    },
  });
  if (state.isMaximized) mainWindow.maximize();
  mainWindow.once('ready-to-show', () => mainWindow.show());
  mainWindow.on('close', () => saveWindowState(mainWindow));
  mainWindow.on('closed', () => { mainWindow = null; });
  mainWindow.webContents.setUserAgent(strippedUA);
  mainWindow.loadURL(START_URL);
}

app.on("open-url", (event, url) => {
  createComposeWindow(url);
});

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on("second-instance", (event, commands, workingDir) => {
    const mailToUrl = commands.find(item => item.startsWith("mailto:"));
    if (mailToUrl) {
      createComposeWindow(mailToUrl);
    } else {
      return;
    }
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

    const mailToUrl = process.argv.find(item => item.startsWith("mailto:"));
    if (mailToUrl) {
      createComposeWindow(mailToUrl);
    } else {
      createWindow(strippedUA);
    }
    
  });

}
