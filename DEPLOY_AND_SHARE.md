# Deploy and Share (No Terminal for End Users)

## Goal

End users should not run `npm`, `build`, or any terminal commands.
They only install a manifest in Excel.

## Recommended flow (GitHub Pages)

1. Push this repo to GitHub.
2. Enable GitHub Pages for the repository.
3. Keep `.github/workflows/deploy-pages.yml` enabled so GitHub Actions builds and publishes `dist/`.
4. Update `manifest.github.xml`:
   - Replace `YOUR_GITHUB_USERNAME` with your username.
   - Keep `/managerial-finance/` path aligned to your repo name if different.
5. Share `manifest.github.xml` (or hosted link) with users.

## End-user install (Windows/Mac Excel)

1. Excel -> Insert -> Get Add-ins -> Upload My Add-in
2. Select shared manifest
3. Open task pane and use calculators

No terminal required for users.

## Optional self-hoster package

- Maintainer can create zip: `npm run package`
- Zip contains prebuilt files + manifest template for teams that want to host internally.
