# Deploy and Share (Offline First, No GitHub Required)

## Goal

End users should run the add-in from downloaded files only:

- no GitHub dependency
- no npm
- no terminal commands

## Offline distribution flow (primary)

1. Maintainer creates release zip (`npm run package`).
2. End user downloads and unzips.
3. End user starts local runtime:
   - Windows: `Start Addin.bat`
   - Mac: `Start Addin.command`
4. End user installs `manifest.offline.xml` in Excel:
   - Insert -> Get Add-ins -> Upload My Add-in
5. Add-in loads from local loopback URL (`127.0.0.1:32123`).

Stop runtime when done:

- Windows: `Stop Addin.bat`
- Mac: `Stop Addin.command`

## What the package includes

- `dist/` prebuilt add-in files
- `runtime/` local Caddy server binaries + config
- `manifest.offline.xml`
- start/stop scripts for Windows and Mac
- docs

## Optional hosted flow (secondary)

If desired, you can still host using GitHub Pages with `manifest.github.xml`, but it is not required for end users.
