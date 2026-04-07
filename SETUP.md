# Setup (TechBI)

## 1) Create SharePoint List

On `https://bekbg.sharepoint.com/sites/TechBI`, create the list and columns exactly as described in:

- `C:\Users\jason.taylor\Downloads\Code\SHAREPOINT-LIST-SETUP.md`

Default list title expected by the web part:

- `CLOCloseoutMatrix`

## 2) Build + Package

From `spfx-closeout-matrix`:

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

## 3) Deploy to SharePoint

1. Upload `sharepoint/solution/closeout-matrix.sppkg` to your App Catalog.
2. Add the app when prompted.
3. Edit a modern page on the `TechBI` site and add the `Closeout Matrix` web part.

## Notes

- If the list name differs, use the web part property pane to change the list title.
- This web part uses standard SharePoint REST via SPFx `SPHttpClient` (no custom hosting required).

