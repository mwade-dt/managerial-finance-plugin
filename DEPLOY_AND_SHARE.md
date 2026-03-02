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
6. Add-in loads from local loopback URL (`127.0.0.1:32123`).

Stop runtime when done:

- Windows: `Stop Addin.bat`
- Windows remove catalog: `Uninstall Addin.bat`
- Mac: `Stop Addin.command`

## What the package includes

- `dist/` prebuilt add-in files
- `runtime/` local Caddy server binaries + config
- `manifest.offline.xml`
- start/stop scripts for Windows and Mac
- install/uninstall catalog scripts for Windows
- docs

## Optional hosted flow (secondary)

If desired, you can still host using GitHub Pages with `manifest.github.xml`, but it is not required for end users.
