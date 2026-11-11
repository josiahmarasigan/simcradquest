# Rad Quest

A gamified staff development program for a radiology department. Staff complete verifiable quests, earn XP and coins, raise an 8-bit pet, and spend coins on real rewards such as training seats.

## How it's put together

```
Phone  ──opens──▶  App shell (GitHub Pages, /docs)
                        │  every action = one HTTPS POST
                        ▼
                   Apps Script web app (/apps-script)
                        │
                        ▼
                   Google Sheet (data) + Drive folder (proof photos)
```

- **App shell** (`/docs`): the screens, pet art and logic. Static files, so GitHub Pages can host them. A service worker keeps them on the phone, so the app opens instantly, works from the home screen, and still shows the last saved pet with no signal.
- **Server** (`/apps-script`): checks who is calling, applies the game rules, and reads and writes the Sheet. Nothing private lives in this repo.

## Repo layout

| Path | What it is |
|---|---|
| `docs/index.html` | The whole app |
| `docs/config.js` | One setting: the server's web app URL |
| `docs/sw.js` | Service worker (offline shell) |
| `docs/manifest.webmanifest`, `docs/icons/` | Makes it installable |
| `apps-script/Code.gs` | Server code |
| `apps-script/appsscript.json` | Apps Script settings |

## Setup

### 1. Server (Google)

1. Create a Google Sheet. Go to **Extensions ▸ Apps Script**.
2. Replace the default `Code.gs` with `apps-script/Code.gs`.
3. Reload the Sheet. Run **Rad Quest ▸ 1. Run setup** and approve the permissions. Note the admin PIN it shows.
4. In Apps Script: **Deploy ▸ New deployment ▸ Web app**. Execute as: **Me**. Who has access: **Anyone**. Copy the **Web app URL** (it ends in `/exec`).

### 2. App shell (GitHub)

1. Paste that URL into `docs/config.js` and commit.
2. In the repo: **Settings ▸ Pages ▸ Deploy from a branch ▸ `main` ▸ `/docs`**.
3. After a minute the app is live at `https://<your-username>.github.io/<repo-name>/`.

### 3. Connect the two

1. In the Sheet's **Config** tab, set `APP_URL` to that GitHub Pages address.
2. Fill in the **Staff** tab.
3. Run **Rad Quest ▸ 2. Generate personal links**, then **3. Email personal links to staff**.

Staff open their link once, then choose **Add to Home Screen**.

## Updating

- **App shell:** commit to `docs/`. Bump `VERSION` at the top of `docs/sw.js` so phones pick up the new files. People see the update on their second open.
- **Server:** paste the new `Code.gs`, then **Deploy ▸ Manage deployments ▸ Edit ▸ Version: New version**. The URL stays the same.

With [clasp](https://github.com/google/clasp) you can push `apps-script/` from the command line instead of pasting: `clasp clone <SCRIPT_ID> --rootDir apps-script`, restore the repo's files with `git checkout -- apps-script`, then `clasp push`.

## Security notes

- Each staff member's link ends in a private code. It works like a password. The app remembers it on the device.
- The server's URL in `config.js` is not a secret: every call must carry a valid personal code, and admin calls also need the PIN from the Config tab.
- Only the actions listed in `api_()` in `Code.gs` can be called from outside.
- **Never commit:** `.clasprc.json`, exports of the Sheet (staff emails, codes, PIN), or proof photos.
- GitHub Pages sites are public even when the repo is private on some plans. The page shows nothing without a personal code, and `noindex` keeps it out of search results.

## Data and privacy

- Evidence must not show patient names, IDs, faces, or PACS and worklist screens. The app asks staff to confirm this on every submission.
- Photos are resized on the phone before upload, which also removes location data.
- The phone keeps a copy of the player's own last state (name, team, coins, pet) so the app can open offline.

## Credits

Concept and program design by the SIMC Radiology staff development lead.
