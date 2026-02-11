# Foundry.Automation

Automation repository responsible for generating and maintaining dynamic download manifests and external resource metadata for Foundry.

## Operating System Catalog

This repository includes a PowerShell generator that pulls Windows catalog metadata from WORProject, downloads CAB catalogs, extracts `products.xml`, and writes deterministic ESD cache outputs.

- Script: `Scripts/Update-OSCatalog.ps1`
- Schema: `Schemas/OperatingSystem.schema.json`
- Outputs:
  - `Cache/OS/OperatingSystem.json`
  - `Cache/OS/OperatingSystem.xml`
  - `Cache/OS/OperatingSystem.md`

Run:

```powershell
pwsh -NoProfile -File ./Scripts/Update-OSCatalog.ps1
```

Prerequisite: `7zz` or `7z` must be available in `PATH` for CAB extraction.
