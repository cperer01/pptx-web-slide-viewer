# Web Slide

A PowerPoint **content add-in** for macOS that embeds a live, interactive
webpage into a slide. Type a URL into the bar at the top of the embedded object,
press Enter, and the page loads in the slide surface below. The URL is saved
with the slide, so reopening the deck restores the page. The point is to flow
from slides → live product demo → slides without ever leaving PowerPoint.

The page renders at a fixed 1920×1080 desktop viewport and is scaled to fit the
slide object.

## How this repo serves the add-in

This repo hosts the add-in's files on **GitHub Pages**, so there's no laptop or
Terminal to keep running during a presentation. The live host is:

```
https://cperer01.github.io/pptx-web-slide-viewer/
```

Every URL in `manifest.xml` already points there. PowerPoint loads the add-in
surface from `…/src/index.html`, and the icons from `…/assets/`.

### One-time setup: turn on GitHub Pages

After this branch is merged to `main`, enable Pages once in the repo UI:

**Settings → Pages → Source: _Deploy from a branch_ → Branch: `main` → Folder:
`/ (root)` → Save.**

Within a minute or two the site goes live at the URL above. The `.nojekyll` file
at the repo root tells Pages to serve the files verbatim. (This toggle can only
be set in the GitHub UI, not via a commit, which is why it's a manual step.)

> GitHub Pages does **not** send `X-Frame-Options`, so the add-in surface frames
> correctly inside the Office host with no extra configuration. Pages also can't
> set custom `Cache-Control` headers, so hosted edits propagate on the next
> deploy (CDN cache, ~10 minutes) rather than instantly.

## The one thing to understand first

This loads pages in an iframe. Any site that refuses to be framed (most
third-party sites) will show a blank frame. It works fully when the target is a
page **you control**, such as your own app, which must send:

```
Content-Security-Policy: frame-ancestors https://*.officeapps.live.com https://*.microsoft.com 'self'
```

and must **not** also send `X-Frame-Options: DENY` or `SAMEORIGIN`.

## Installation

The hosted files being public does not by itself put the add-in on anyone's
PowerPoint. To install on a Mac:

1. Download **`install.command`** and **`manifest.xml`** (keep them together in
   the same folder, e.g. `~/Downloads`).
2. In Terminal, run `chmod +x ~/Downloads/install.command`. Accept any popups.
3. In Terminal, run `xattr -c ~/Downloads/install.command`. Accept any popups.
4. Double-click **`install.command`**. Accept any popups.
5. Open your presentation, then select **Add-ins → Developer Add-ins → Web
   Slide**.

> Steps 2–3 clear the executable bit and the macOS download-quarantine flag so
> the installer runs without the "unidentified developer" block. Keep
> `manifest.xml` next to `install.command` — the installer copies the manifest
> sitting beside it.

To send it to a non-technical teammate, zip `manifest.xml` + `install.command`
together; everything else is served from GitHub Pages.

### Team-wide rollout (zero-touch)

An admin can upload `manifest.xml` once in the Microsoft 365 admin center under
**Integrated Apps (Centralized Deployment)** and assign it to users or groups.
It then appears automatically in their PowerPoint on Mac, Windows, and the web,
with nothing for them to install. Propagation can take up to 24 hours.

## Local development (optional)

You only need this to edit the add-in surface against a local server instead of
the hosted files:

```bash
npm run setup-certs   # once: trusts a localhost cert via office-addin-dev-certs
npm start             # serves this folder at https://localhost:3000
```

Then point `manifest.xml` back at `https://localhost:3000` (swap it for the
GitHub Pages origin) and install as above. The dev server (`server.js`) sends
`Cache-Control: no-store`, so edits to `src/` show on reopening the deck without
clearing the Office cache.

## Tuning

- The embedded page renders at 1920×1080 and scales to fit. To hide a fixed
  header or button bar inside the embedded app, set `CROP_TOP` in `src/app.js`
  to roughly its pixel height at 1920 wide (for example 72 or 96).
- Edit mode is fully interactive. Slideshow-mode interactivity is reliable on
  Windows and generally works on Mac, but test your exact build before relying
  on clicking inside the frame mid-present, and keep a fallback slide.

## Notes

- The manifest is the add-in-only XML format on purpose. The unified JSON
  manifest does not sideload on Mac, so do not convert it.
- To remove the add-in, delete `manifest.xml` from
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` and clear
  the Office cache.

## Repo layout

| Path | Purpose |
| --- | --- |
| `src/` | The add-in surface PowerPoint loads (`index.html`, `app.js`, `styles.css`, `config.js`). |
| `assets/` | Add-in icons referenced by `manifest.xml`. |
| `manifest.xml` | The add-in manifest; install this into PowerPoint. |
| `install.command` | Mac double-click installer for `manifest.xml`. |
| `server.js`, `package.json` | Optional local HTTPS dev server. |
| `.nojekyll` | Tells GitHub Pages to serve files verbatim. |
