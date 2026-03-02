# Deploy and Share (Offline First, No GitHub Required)

## Goal

End users should run the add-in from downloaded files only:

- no GitHub dependency
- no npm
- no terminal commands

## Offline distribution flow (primary)

1. Maintainer creates release zip (`npm run package`).
2. End user downloads and unzips.
3. Windows one-time setup: run `Install Addin.bat` to register the trusted add-in catalog.
4. End user starts local runtime:
   - Windows: `Start Addin.bat`
   - Mac: `Start Addin.command`
5. End user adds the add-in in Excel:
   - Home -> Add-ins -> More Add-ins -> My Add-ins -> SHARED FOLDER
   - Select "Managerial Finance Tools (Offline)" and click Add
   - Do not use "Upload My Add-in" for Windows desktop; many builds do not show it.
6. Add-in loads from local loopback URL (`127.0.0.1:32123`).

## Stop runtime when done

- Windows: `Stop Addin.bat`
- Windows remove catalog: `Uninstall Addin.bat`
- Mac: `Stop Addin.command`

## Troubleshooting (Windows)

- If `SHARED FOLDER` does not appear, close Excel, rerun `Install Addin.bat`, then reopen Excel.
- If the add-in is listed but does not load, make sure `Start Addin.bat` is running first.
- If the extracted folder is moved, run `Uninstall Addin.bat` and then `Install Addin.bat` from the new location.

## What the package includes

- `dist/` prebuilt add-in files
- `runtime/` local Caddy server binaries + config
- `manifest.offline.xml`
- launcher scripts for install/start/stop/remove on Windows, plus start/stop on Mac
- docs

## Optional hosted flow (secondary)

If desired, you can still host using GitHub Pages with `manifest.github.xml`, but it is not required for end users.
