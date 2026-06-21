# Liquidity Tracker — setup guide

A small installable app (PWA) that shows your monthly recurring payments and reads/writes the same
workbook in your OneDrive (the one with the `AppLog` / `AppData` sheets). It works on Android (installed to
your home screen) and in any browser on your Mac. Both the app and Excel edit the one spreadsheet, so they
always agree.

There is no server and no database — the app talks only to **your** OneDrive through Microsoft's official
Graph API. Nothing is sent anywhere else.

> **Where these files live:** the app code sits in `~/Claude/Projects/MyFinanceApp` (a normal local folder,
> deliberately **outside** OneDrive so git and OneDrive don't conflict). It is a git repository linked to
> `github.com/rhkoncloud/liquidity-tracker`. Your **workbook stays in OneDrive** — the app reaches it over
> the network, so it doesn't matter that the code and the data live in different places.
>
> **Privacy note:** the exact path to your workbook is **not** stored in the code or this repo. You enter it
> once inside the app (Settings ⚙) and it's saved locally on each device — so this public repository never
> reveals where your financial data lives.

Total setup time: about 15 minutes, done once.

---

## What you'll do

1. Register the app with Microsoft (free) to get a Client ID.
2. Put the app's files online over HTTPS (GitHub Pages — free).
3. Paste your Client ID into `config.js`.
4. Open the page on your phone, sign in, and enter your workbook path once (Settings ⚙); then "Add to Home screen".

---

## Step 1 — Register the app with Microsoft (get a Client ID)

### Step 1a — Get a directory (one-time, only if prompted)

If you sign in with a **personal Microsoft account** (e.g. `…@outlook.com`) you may see a warning like:
*"These applications are associated with the account … but are not contained within any directory. The
ability to create applications outside of a directory has been deprecated…"* with two suggested links.

**Choose "signing up for Azure"** — not the M365 Developer Program. Here's why:

- A **free Azure account automatically creates a free Microsoft Entra directory** for you, which is exactly
  the container the app registration now needs. The Entra ID Free tier covers app registration at no cost
  (a credit card is requested only to verify identity and is **not** charged).
- The **M365 Developer Program is the wrong fit**: it's now restricted (mostly requires a Visual Studio
  subscription, and new accounts are often rejected), and even if granted it gives a *separate sandbox
  tenant* with fake test users and an empty OneDrive — it would **not** contain your real `Finances`
  spreadsheet.

So: click **"signing up for Azure"**, complete the free signup (it provisions your directory), then continue
with Step 1b below. A fresh directory can take a few minutes to appear — if "New registration" misbehaves,
refresh or re-open the App registrations page.

> The directory is only *where the registration is stored*. Your real OneDrive is still reached through your
> personal account, because in Step 1b you set the account type to include "personal Microsoft accounts".

### Step 1b — Register the app

1. Go to **https://entra.microsoft.com** (or portal.azure.com → "Microsoft Entra ID") and sign in with the
   same Microsoft account that owns this OneDrive.
2. In the left menu: **App registrations → + New registration**.
3. Name: `Liquidity Tracker`.
4. Supported account types: choose **"Accounts in any organizational directory and personal Microsoft
   accounts"** (this is required so your personal OneDrive account can sign in).
5. Redirect URI: set the platform dropdown to **Single-page application (SPA)** and enter the URL where
   you'll host the app (from Step 2). If you don't know it yet, leave it blank and add it after Step 2.
6. Click **Register**.
7. On the app's Overview page, copy the **Application (client) ID** — a string like
   `1234abcd-….` You'll paste it into `config.js` in Step 3.

API permissions are already correct by default: the app requests `User.Read` and `Files.ReadWrite` at
sign-in, which Microsoft grants for your own files. You do **not** need to add a client secret — a
single-page app must never hold one, and this app doesn't.

> Note on `Files.ReadWrite`: Microsoft's delegated permission model grants this app access to your files
> on your behalf. There is no built-in "just this one file" scope for arbitrary OneDrive files, so the
> permission is broad even though the app only ever touches the one workbook in its code. If you'd like to
> limit blast radius, you can keep the workbook in a dedicated folder and/or sign out when not using the app.

---

## Step 2 — Host the files over HTTPS (GitHub Pages)

A PWA must be served over HTTPS (this is also what keeps it secure). GitHub Pages is the simplest free option.

The repository is **already created and linked**: the local folder `~/Claude/Projects/MyFinanceApp` is a git
repo whose `origin` points to `github.com/rhkoncloud/liquidity-tracker` (public).

1. **Push the code to GitHub** (from a Terminal):
   ```
   cd ~/Claude/Projects/MyFinanceApp
   git push -u origin main
   ```
   The first push asks you to sign in to GitHub (a browser window opens, or your Mac's credential helper
   handles it). If it doesn't prompt cleanly, install the GitHub CLI once — `brew install gh` then
   `gh auth login` — and re-run the push. (GitHub Desktop also works: **Add local repository → choose this
   folder → Push origin**.)
2. In the repo on GitHub: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, branch
   `main`, folder `/ (root)`, **Save**.
3. After a minute, GitHub shows your live URL, e.g. `https://rhkoncloud.github.io/liquidity-tracker/`.
4. Go back to your Azure app registration (Step 1) → **Authentication** → add that exact URL as a
   **Single-page application** redirect URI (include the trailing slash). Save.

> A public repo only exposes the app's code (which contains no secrets) — never your spreadsheet or your
> data. Your financial data stays in OneDrive and is only fetched into the app after you sign in.

**Making changes later:** edit the files in `~/Claude/Projects/MyFinanceApp`, then `git add -A && git commit
-m "…" && git push`. GitHub Pages redeploys automatically within a minute.

---

## Step 3 — Add your Client ID

`config.js` holds only your **public Client ID** — no secrets, and (deliberately) **not** your file path.

1. Open `config.js`.
2. Replace `PASTE_YOUR_CLIENT_ID_HERE` with the Application (client) ID from Step 1.
3. Leave `redirectUri` as `null` (it auto-uses the page URL) unless you host under a sub-path and want to
   pin it explicitly.
4. Save, then publish the change: `cd ~/Claude/Projects/MyFinanceApp && git add config.js && git commit -m
   "Add client id" && git push`. (Or use GitHub Desktop.)

> Your workbook's path is **not** put here. You'll enter it once inside the app after signing in (Step 4),
> and it's saved locally on your device — keeping it out of the public repo.

---

## Step 4 — Install on your phone

1. On your Android phone, open the GitHub Pages URL in **Chrome**.
2. Tap **Sign in with Microsoft**, approve the permission prompt once.
3. **Enter your workbook path** when prompted (or via Settings ⚙): the path to the workbook inside your
   OneDrive, relative to the OneDrive root, e.g. `Folder/Subfolder/YourWorkbook.xlsx` — no leading slash.
   This is saved on the device only. Repeat once on each device you use (phone, Mac).
4. Chrome menu (⋮) → **Add to Home screen / Install app**. An icon appears like a normal app; it opens
   full-screen.
5. On your Mac, open the same URL in any browser, sign in and enter the path once, or in Chrome use
   **Install** from the address bar.

That's it. Pick a month, tap a row to mark it paid, tap ✎ to change an amount, or **+ Add expense** to add
one. Every change is appended to the `AppLog` sheet and the `AppData` snapshot is refreshed, so your Excel
monthly summary stays current.

---

## How your data is kept secure

Built into the app:

- **No backend, no third parties.** The app only ever calls `graph.microsoft.com` (your OneDrive) and
  Microsoft's login servers. There is no analytics, tracking, or external server that could see your data.
- **Modern sign-in (OAuth 2.0 auth-code + PKCE)** via Microsoft's official MSAL library. The app never sees
  or stores your password, and holds **no client secret**.
- **Strict Content-Security-Policy.** The page is locked down so it can only load its own code plus the
  Microsoft login library, and can only connect to Microsoft endpoints. This blocks the common script-
  injection (XSS) attacks that would otherwise be the main risk for a data app.
- **All app logic is in a separate file** (`app.js`) with no inline scripts, which is what lets the CSP be
  strict.
- **The offline cache never stores your data.** The service worker caches only the app's static files (the
  UI shell). It is written to deliberately ignore all Graph and login traffic, so your figures and access
  tokens are never written to disk by the app.
- **HTTPS only** (enforced by GitHub Pages and required for the service worker).

What you should do:

- **Put a screen lock on your phone** (PIN/biometric). That is the real protection for an installed app —
  it keeps anyone who picks up your phone out of the app.
- **Sign out** (button top-right) on any shared or public computer.
- **Keep the repo public-code / private-data distinction in mind**: never paste anything other than the
  Client ID into `config.js`, and never commit the spreadsheet to the repo.
- **Optional hardening — pin the MSAL library:** the app loads MSAL from jsDelivr
  (`cdn.jsdelivr.net/npm/@azure/msal-browser@2/lib/msal-browser.min.js`). For extra safety you can pin an
  exact version and add its Subresource Integrity hash — jsDelivr shows a "copy SRI" button next to each
  file — to the `<script src="…msal-browser.min.js" …>` tag in `index.html`. This makes the browser reject
  the library if it were ever tampered with. (Left off by default because a wrong hash silently blocks the app.)
- **Review app access anytime** at https://account.microsoft.com → Privacy / Apps, where you can revoke the
  app's access with one click.

---

## How edits from two devices are reconciled (conflict handling)

The app uses an **append-only event log** so edits from your Mac and phone can never overwrite each other —
even if they're made offline and arrive out of order.

- Every change you make (mark paid, edit an amount, add or remove an item) is **appended** as a timestamped
  event to the `AppLog` sheet (an Excel table called `tblLog`). Appending never overwrites existing data, so
  two devices writing at once just add two rows — nothing is lost.
- The current state you see is **rebuilt by replaying those events in timestamp order**. Because ordering is
  by the timestamp stamped at the moment of the edit — not by when the write reaches OneDrive — a late-
  arriving older edit can't clobber a newer one. If two devices change the very same thing, the one made
  later (by the clock) wins, deterministically.
- The `AppData` sheet is a **derived snapshot** the app refreshes for the human-readable monthly summary. It
  is not the source of truth — if it ever drifts, opening the app rebuilds it from the log. `AppLog` is the
  record of truth and doubles as a full audit trail of every change.
- Each device gets a random ID stored locally, so the log also shows which device made each edit.

One thing still worth a habit: the **file-level** lock. If you have the workbook open in **desktop Excel**
while the app writes, OneDrive can occasionally create a "conflicted copy" of the whole file (this is
OneDrive's behaviour, separate from the app). The app edits cells in place rather than rewriting the file,
which merges far better, but the safe routine is to treat the app as the primary editor and not leave the
file open in Excel during heavy app use.

## Good to know / limitations

- **Adding an amount to a month an item didn't previously have** is done via "+ Add expense" (it creates the
  item across the year). Editing an existing row changes that month's amount.
- **Amounts shown are last year's actuals +8%, plus the items you've told me about** (school fees, badminton,
  StemCyte, advance tax). Adjust any of them in the app or in Excel.

---

## Files in this folder

| File | Purpose |
|------|---------|
| `index.html` | The app's markup + security policy |
| `app.js` | All app logic (sign-in, read/write, UI) |
| `config.js` | **Your** public Client ID (the only file you edit; no path, no secrets) |
| `manifest.webmanifest` | Makes it installable as an app |
| `sw.js` | Offline support for the UI (never caches data) |
| `icon-192.png`, `icon-512.png` | App icons |
| `SETUP_GUIDE.md` | This guide |

In your workbook, the app uses two sheets: **`AppLog`** (the append-only event log / source of truth and
audit trail) and **`AppData`** (a derived monthly snapshot for easy reading in Excel). Your original
analysis sheets and the `Liquidity Calendar` are untouched.
