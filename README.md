# outlook-for-linux

Unofficial Electron wrapper for Microsoft Outlook Web (`https://outlook.office.com/mail`).
Microsoft ships no Linux desktop client; this wraps OWA in a native window with
persistent login, web notifications, and one hard rule:

**Every clicked link (http/https/mailto) opens in the system default browser — never inside the app.**

## How it works

- `src/url-policy.js` — pure allowlist module (runs under bare node, no Electron):
  - `isInternalNav`: only `https:` on known Microsoft mail/auth hosts renders in-app.
  - `isOpenableExternal`: only `http:`, `https:`, `mailto:` are ever forwarded to the OS.
- `src/main.js` attaches both handlers at a single choke point (`app.on('web-contents-created')`):
  - `setWindowOpenHandler` denies all popups and routes the URL out (middle-click, ctrl+click, `target="_blank"`, `window.open`).
  - `will-navigate` blocks same-window navigation to non-internal hosts and routes it out instead.
- Session uses partition `persist:outlook` (login survives restarts), sandboxed renderer,
  no Node in the page, Electron marker stripped from the user agent.

## Run

```sh
npm ci
npm start
```

## Test

```sh
npm test        # url-policy unit asserts (plain node)
```

## Install / package

Reference PKGBUILD in `packaging/`. It installs app files to `/usr/lib/outlook-for-linux`,
a `/usr/bin/outlook-for-linux` launcher using system `electron`, plus `.desktop`
entry and hicolor icons.

## Install (package from source)
```sh
tar cvf outlook-for-linux-*version*.tar src assets packaging LICENSE package.json README.md
mv outlook-for-linux-*version*.tar packaging/
cd packaging
makepkg .
sudo pacman -U outlook-for-linux-*version*-1-any.pkg.tar.zst
```

## Install (user-local, from this checkout)

```sh
install -m 755 packaging/outlook-for-linux.sh ~/.local/bin/outlook-for-linux
cp packaging/outlook-for-linux.desktop ~/.local/share/applications/
for s in 16 32 48 64 128 256 512; do
  install -Dm644 assets/icons/outlook-for-linux-$s.png \
    ~/.local/share/icons/hicolor/${s}x${s}/apps/outlook-for-linux.png
done
install -Dm644 assets/outlook-for-linux.svg \
  ~/.local/share/icons/hicolor/scalable/apps/outlook-for-linux.svg
update-desktop-database ~/.local/share/applications
gtk-update-icon-cache -tf ~/.local/share/icons/hicolor
```

Icon source: official Outlook product mark from Microsoft's Office Fabric
brand-icons CDN (`static2.sharepointonline.com/.../brand-icons/product/svg/outlook_48x1.svg`),
rendered to the hicolor size set with rsvg-convert. Microsoft trademark — personal use.

## Notes

- Window state (bounds, maximized) persists to `$XDG_CONFIG_HOME/outlook-for-linux/window-state.json`.
- Permissions granted to the page: notifications, media (camera/mic), fullscreen,
  sanitized clipboard write. Everything else denied.

## Troubleshooting

Two stderr messages at startup are known-benign on Wayland and safe to ignore:

- `'--ozone-platform=wayland' is not compatible with Vulkan` — Electron spawns its
  GPU process via zygote without the Wayland ozone flags (electron/electron#50455);
  Chromium logs this and falls back to GL compositing. Not app-controllable;
  no command-line switch silences it.
- `vaInitialize failed: unknown libva error` — no VA-API driver for your GPU.
  On Intel Gen12+ (Alder Lake and newer) install `intel-media-driver` to get
  hardware video decode; without it Chromium software-decodes video.
