# MyCloudPlay

A static, browser-based OneDrive music player. Track lists stay sorted by filename. The original cloud logo and animated playback bars are retained.

## Run

From this directory:

```sh
python3 -m http.server 8000 --bind 127.0.0.1
```

The checked-in `styles.css` is ready to serve; running the site does not require Node.js or an application server.

OneDrive uses the existing Microsoft application and `Files.ReadWrite.All` permission. Its registered redirect URI must match the served origin plus `/`. No client secret is required by this browser application. Production should use HTTPS.

## Edit styles

Node.js 20 or newer is recommended. Install the locked development dependencies and regenerate the stylesheet after editing HTML, JavaScript class names, or CSS:

```sh
npm ci
npm run build
```

`npm run dev:css` watches files while developing. Commit the generated `styles.css` with its source changes. No Tailwind runtime CDN is used.

- `index.html`: markup, original logo artwork, English interface copy.
- `appearance.js`: safe browser storage, five color skins, Light/Dark/System modes; applied before paint.
- `app.js`: playback, Microsoft sign-in, OneDrive file operations, accessible dialogs.
- `styles/app.css`: Tailwind input and appearance tokens.
- `styles.css`: generated stylesheet served to browsers.

The appearance picker is in the header. Ocean, Moss, Sand, Rose, and Graphite each support light and dark modes. The existing `theme` preference is preserved and the new `skin` preference is saved in the same browser.

## Production deployment

Vercel builds the `main` branch using `vercel.json`. `npm run build` regenerates the stylesheet and copies only the six public runtime files into `dist/`. Source styles, tests, development dependencies, and documentation are not included in the served deployment.

## Browser tests

Python 3.10 or newer:

```sh
python3 -m pip install -r requirements-dev.txt
python3 -m playwright install chromium
npm test
```

Tests automatically use a system `chromium` when available, or Playwright's installed Chromium. Set `CHROMIUM_EXECUTABLE` to select another Chromium executable. The tests start their own local server.

The suite uses fixture responses for Microsoft sign-in, OneDrive, and audio, and makes no changes to any real account. It covers responsive controls, appearance persistence, keyboard dialogs, corrupt/disabled storage, playback state, sample history, safe metadata rendering, paginated filename sorting, stale playback requests, and file-operation success/failure/cancellation. Real Microsoft sign-in, account policies, and OneDrive upload/streaming still require an authenticated integration check.

Google Analytics and the existing browser storage/sign-in approach are unchanged. About explains the current data handling and file permissions. The existing `sw.js` is not registered by this revision; offline music is not supported.
