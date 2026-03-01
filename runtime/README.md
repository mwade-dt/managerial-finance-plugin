# Runtime bundle

This folder contains the local static server runtime used by offline installs.

- `bin/windows_amd64/caddy.exe`
- `bin/mac_amd64/caddy`
- `bin/mac_arm64/caddy`
- `Caddyfile` (serves `../dist` on `127.0.0.1:32123`)

End users launch via:

- `Start Addin.bat` (Windows)
- `Start Addin.command` (Mac)

No Node/npm runtime is required for end users.
