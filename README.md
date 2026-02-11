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

## Dell Catalog

This repository also includes a PowerShell generator that pulls Dell Driver Pack Catalog metadata from a CAB, extracts `DriverPackCatalog.xml`, and writes split outputs for DriverPack and WinPE catalogs.

- Script: `Scripts/Update-DellCatalog.ps1`
- Schema: `Schemas/DellCatalog.schema.json`
- Outputs:
  - `Cache/DriverPack/Dell/DriverPack_Dell.json`
  - `Cache/DriverPack/Dell/DriverPack_Dell.xml`
  - `Cache/DriverPack/Dell/DriverPack_Dell.md`
  - `Cache/WinPE/Dell/WinPE_Dell.json`
  - `Cache/WinPE/Dell/WinPE_Dell.xml`
  - `Cache/WinPE/Dell/WinPE_Dell.md`

Run:

```powershell
pwsh -NoProfile -File ./Scripts/Update-DellCatalog.ps1
```
