# Rad Quest

A gamified staff development program for a radiology department. Staff complete verifiable quests and quizzes, earn XP and coins, raise an 8-bit pet, and spend coins on real rewards such as training seats.

## How it's put together

```
Phone ──opens──▶ Loader (GitHub Pages, /docs)          set up once, never changes
                    │  1 HTTPS call: "boot"
                    ▼
                 Apps Script web app (/apps-script)     everything you update lives here
                    ├─ App.html  the screens
                    └─ Code.gs   rules, quizzes, admin
                    │
                    ▼
                 Google Sheet (data) + Drive folder (proof photos)
```

- **Loader** (`/docs`): a tiny page that asks Apps Script for the latest screens and the player's data, then shows them. It keeps a copy on the phone so the app opens with no signal, and makes the app installable (home-screen icon, full screen).
- **Apps Script** (`/apps-script`): `App.html` (every screen) and `Code.gs` (the server). **Deploy a new version in Apps Script and every phone gets it on its next open.** No GitHub upload needed.

## Repo layout

| Path | What it is | How often it changes |
|---|---|---|
| `apps-script/App.html` | All the app's screens | Whenever you improve the app |
| `apps-script/Code.gs` | Server code, quiz bank | Whenever you improve the app |
| `apps-script/appsscript.json` | Apps Script settings | Rarely |
| `docs/index.html` | Loader | Never |
| `docs/config.js` | The server's web app URL | Once |
| `docs/sw.js`, `docs/manifest.webmanifest`, `docs/icons/` | Offline + install | Never |

## Setup (once)

### 1. Apps Script

1. Create a Google Sheet. Go to **Extensions ▸ Apps Script**.
2. Replace the default `Code.gs` with `apps-script/Code.gs`.
3. Click **+ ▸ HTML**, name it exactly `App`, and paste in `apps-script/App.html`.
4. Reload the Sheet. Run **Rad Quest ▸ 1. Run setup** and approve the permissions. Note the admin PIN it shows.
5. **Deploy ▸ New deployment ▸ Web app**. Execute as: **Me**. Who has access: **Anyone**. Copy the **Web app URL** (ends in `/exec`).

### 2. GitHub

1. Paste that URL into `docs/config.js` and commit.
2. **Settings ▸ Pages ▸ Deploy from a branch ▸ `main` ▸ `/docs`**.
3. After a minute the app is live at `https://<your-username>.github.io/<repo-name>/`.

### 3. Connect

1. In the Sheet's **Config** tab, set `APP_URL` to the GitHub Pages address.
2. Fill in the **Staff** tab.
3. Run **Rad Quest ▸ 2. Generate personal links**, then **3. Email personal links to staff**.

Staff open their link once and choose **Add to Home Screen**.

## Updating

1. In Apps Script, change `App.html` and/or `Code.gs`.
2. **Deploy ▸ Manage deployments ▸ Edit ▸ Version: New version ▸ Deploy.** The URL stays the same.

That's it. Each phone checks a fingerprint of `App.html` when it opens and downloads the screens only when they've changed. Copy your changes back into this repo when you want a backup.

With [clasp](https://github.com/google/clasp) you can push `apps-script/` from the command line: `clasp clone <SCRIPT_ID> --rootDir apps-script`, restore the repo's files with `git checkout -- apps-script`, then `clasp push`.

## Quizzes

Nine starter quizzes (121 questions): X-ray, CT and ultrasound (technical and knowledge), radiation protection, physics, and SIMC policies (Draft 0). Answers stay on the server. Passing pays coins and XP at once. Edit questions in the **Questions** tab; results are in **QuizAttempts**; `QUIZ_TRIES_PER_DAY` in Config limits retries.

## Security notes

- Each staff member's link ends in a private code. It works like a password. The app remembers it on the device.
- The URL in `config.js` is not a secret: every call needs a valid personal code, and admin calls also need the PIN from the Config tab.
- Only the actions listed in `api_()` in `Code.gs` can be called from outside.
- **Never commit:** `.clasprc.json`, exports of the Sheet (staff emails, codes, PIN), or proof photos.
- The GitHub Pages site is public. It shows nothing without a personal code and is hidden from search engines.

## Data and privacy

- Evidence must not show patient names, IDs, faces, or PACS and worklist screens. The app asks staff to confirm this on every submission.
- Photos are resized on the phone before upload, which also removes location data.
- The phone keeps a copy of the screens and of the player's own last state so the app can open offline.

## Credits

Concept and program design by the SIMC Radiology staff development lead.
