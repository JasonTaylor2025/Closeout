# Provision SharePoint List (No PowerShell)

This creates the `CLOCloseoutMatrix` list and all required columns on:

- `https://bekbg.sharepoint.com/sites/TechBI`

It runs in the browser using your current SharePoint sign-in.

## How To Run

1. Open the site in your browser: `https://bekbg.sharepoint.com/sites/TechBI`
2. Press `F12` to open DevTools.
3. Open the **Console** tab.
4. Copy/paste the contents of `provision-sharepoint-list.js` and press Enter.
5. It will log progress and end with `Done`.

## Notes

- You must have permission to create lists/fields on the site.
- If the list already exists, the script will keep it and only create missing fields.
- The SPFx web part expects these internal field names, so this is safer than creating columns manually.

