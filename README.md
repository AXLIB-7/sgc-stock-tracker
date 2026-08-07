# SGC Stock Tracker

A single-page inventory tracker for Seattle Gummy Co. Signs in with a Microsoft
work account, reads/writes inventory + take-out log data as a JSON file in the
signed-in user's OneDrive (via Microsoft Graph), and can export the take-out
log as CSV or the stock levels as an Excel workbook.

## Files

```
index.html               markup only
styles.css                all styling
app.js                     all app logic (auth, OneDrive sync, rendering)
staticwebapp.config.json  Azure Static Web Apps routing / headers config
```

No build step, no npm install, no bundler — these are plain static files.
Open `index.html` in a browser (or serve the folder with any static file
server) and it works. That also means Azure Static Web Apps can deploy it
with **App location: `/`**, **Output location:** *(leave blank)*, and
**Build command:** *(leave blank)* — there is nothing to build.

Two third-party libraries are loaded from a CDN at runtime (not vendored
locally): `xlsx` for the Excel export, and `@azure/msal-browser` for
Microsoft sign-in.

## Running locally

Any static file server works, e.g.:

```
npx serve .
```

or the Azure Static Web Apps CLI, which also proxies auth correctly:

```
npm install -g @azure/static-web-apps-cli
swa start .
```

## Deploying to Azure Static Web Apps

1. In the Azure Portal, create a **Static Web App** resource and point it at
   this repo/folder using the GitHub deployment option. Azure will create a
   `.github/workflows/azure-static-web-apps-*.yml` file in your repo
   automatically — you don't need to write one by hand.
2. When asked for build details, use:
   - **App location:** `/` (or wherever this folder sits in your repo)
   - **Output location:** *(leave empty)*
3. Once deployed, note the app's URL (something like
   `https://<random-name>.azurestaticapps.net`), or set up a custom domain.

## ⚠️ Required: update the Azure AD redirect URI

This app authenticates with `MSAL.js` against an Azure AD (Entra ID) app
registration (`CLIENT_ID` / `TENANT_ID` in `app.js`). Azure AD only allows
sign-in redirects to **exact URLs you've explicitly allow-listed**.

`app.js` computes the redirect URI at runtime from wherever it's actually
being served —

```js
const REDIRECT_URI = window.location.origin + window.location.pathname;
```

— so the same code works on GitHub Pages, Azure Static Web Apps, a custom
domain, or localhost without editing it. But whichever URL that resolves to
in production **must be added** to the app registration:

1. **Entra ID** → **App registrations** → your app → **Authentication**
2. Under **Redirect URIs**, add the exact production URL, e.g.
   `https://<your-app>.azurestaticapps.net/`
3. Save.

You can keep the old GitHub Pages URL registered alongside the new Azure one
if you want both to keep working during the transition — Azure AD allows
multiple redirect URIs on one app registration.

## Data storage

Inventory and take-out log data are stored as a single JSON file
(`inventory.json`) inside a `SGC Stock Tracker` folder in the signed-in
user's own OneDrive — there's no separate database or backend. Moving hosts
doesn't affect this data at all, since it lives in OneDrive, not on the
server.
