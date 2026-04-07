# Closeout Matrix (SPFx Web Part)

This is an SPFx (React) web part version of the Closeout Matrix app. It reads and writes shared data to a SharePoint List.

Target site: `https://bekbg.sharepoint.com/sites/TechBI`  
Target list title: `CLOCloseoutMatrix`

## Build

From `spfx-closeout-matrix`:

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` will be created under `sharepoint/solution/`.

## Deploy

1. Upload the `.sppkg` to your App Catalog (tenant or site collection).
2. Approve permissions if prompted.
3. Add the web part to a modern SharePoint page on the `TechBI` site.

## SharePoint List

Create the list and columns exactly as documented in `C:\Users\jason.taylor\Downloads\Code\SHAREPOINT-LIST-SETUP.md`.

