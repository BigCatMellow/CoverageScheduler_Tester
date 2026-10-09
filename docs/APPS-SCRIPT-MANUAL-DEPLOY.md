# Deploying to Apps Script without `clasp`

The maintainable source stays split across `google-apps-script/`, but you do **not** need to recreate all of those files manually in Apps Script.

GitHub Actions builds a copy-ready artifact after every push to `main`.

## Normal update workflow

1. Open the repository's **Actions** tab.
2. Open the latest **Build Apps Script Deploy Bundle** run.
3. Download the artifact named **coverage-scheduler-apps-script**.
4. Unzip it.
5. In the Apps Script tester project, replace:
   - `CoverageScheduler.gs`
   - `index.html` (the Apps Script HTML file is named `index`)
   - `appsscript.json`
6. Save.
7. Go to **Deploy → Manage deployments → Edit → New version → Deploy**.

## First time switching to the bundle

The bundle combines all server-side modules into `CoverageScheduler.gs` and inlines the web UI partials into `index.html`.

After adding the bundled files, remove the old individual server `.gs` modules and the UI partial HTML files from the Apps Script project. Otherwise Apps Script will see duplicate function/constant definitions.

The optional legacy sidebar files are not required by the full-page web app.

## Local build, if ever needed

```
node scripts/build-apps-script-deploy.js
```

This creates `apps-script-deploy/` with the same files uploaded by GitHub Actions.
