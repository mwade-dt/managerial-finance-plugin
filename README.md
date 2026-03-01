# Managerial Finance Excel Add-in (Offline Office Add-in)

This project ships a downloadable, offline-friendly Excel Office Add-in for managerial finance problems.

End users do not need GitHub, npm, or terminal commands.

## Direct end-user install (downloaded project only)

1. Download and unzip the project (or release zip).
2. Start the local runtime server:
   - Windows: from the unzipped folder, double-click `Start Addin.bat` (or right-click -> Run)
   - Mac: double-click `Start Addin.command`
3. Open Excel and install the offline manifest:
   - Windows (Microsoft 365 desktop): Home -> Add-ins -> More Add-ins -> My Add-ins -> Upload My Add-in
   - If your Excel build shows the older menu: Insert -> Get Add-ins -> My Add-ins -> Upload My Add-in
   - Choose `manifest.offline.xml`
4. Open the add-in and use calculators.
5. Optional stop:
   - Windows: `Stop Addin.bat`
   - Mac: `Stop Addin.command`

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
