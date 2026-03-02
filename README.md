# Managerial Finance Excel Add-in (Offline Office Add-in)

This project ships a downloadable, offline-friendly Excel Office Add-in for managerial finance problems.

End users do not need GitHub, npm, or terminal commands.

## Direct end-user install (downloaded project only)

1. Download and unzip the project (or release zip).
2. Windows only (one-time setup): double-click `Install Addin.bat`.
   - This registers a trusted add-in catalog for this folder in your user profile.
3. Start the local runtime server:
   - Windows: double-click `Start Addin.bat`
   - Mac: double-click `Start Addin.command`
4. Open Excel and add the add-in:
   - Home -> Add-ins -> More Add-ins -> My Add-ins -> SHARED FOLDER
   - Select "Managerial Finance Tools (Offline)" and click Add
5. Open the add-in and use calculators.
6. Optional stop/remove:
   - Windows stop server: `Stop Addin.bat`
   - Windows remove catalog: `Uninstall Addin.bat`
   - Mac stop server: `Stop Addin.command`

## Troubleshooting (Windows)

- If `SHARED FOLDER` does not appear, close Excel, run `Install Addin.bat` again, then reopen Excel.
- If the add-in is listed but does not load, make sure `Start Addin.bat` is running first.

## What is included

- `dist/` prebuilt task pane bundle
- `runtime/` portable local server binaries and config
- `manifest.offline.xml` local-loopback manifest
- one-click launcher scripts for start/stop

## Topics implemented

- TVM (lump sum)
- NPV / IRR
- MIRR
- Payback / Discounted Payback
- EAC
- WACC
- DCF (FCF + terminal growth)
- Relative valuation (multiple x metric)
- Risk/return (one asset)
- Two-asset portfolio
- N-asset portfolio
- Efficient frontier (2-asset grid)
- CAPM
- Bond price / YTM
- Gordon growth stock value
- Core financial ratios

## Maintainer commands (not for end users)

- `npm run build` - rebuild `dist/`
- `npm run package` - create offline release zip in `release/`

## Additional docs

- `DEPLOY_AND_SHARE.md` - offline distribution details
- `TEST_CASES.md` - calculation validation notes
