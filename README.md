# Foundry.Automation

Automation repository responsible for generating and maintaining dynamic download manifests and external resource metadata for Foundry.

## Operating System Catalog

This repository includes a PowerShell generator that pulls Windows catalog metadata from WORProject, downloads CAB catalogs, extracts `products.xml`, and writes deterministic ESD cache outputs.

- Script: `Scripts/Update-OSCatalog.ps1`
- Schema: `Schemas/OperatingSystem.schema.json`
- Outputs:
  - `Cache/OS/OperatingSystem.json`
  - `Cache/OS/OperatingSystem.xml`
  - `Cache/OS/README.md`

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
  - `Cache/DriverPack/Dell/README.md`
  - `Cache/WinPE/Dell/WinPE_Dell.json`
  - `Cache/WinPE/Dell/WinPE_Dell.xml`
  - `Cache/WinPE/Dell/README.md`

Run:

```powershell
pwsh -NoProfile -File ./Scripts/Update-DellCatalog.ps1
```

## HP Catalog

This repository includes a PowerShell generator that pulls HP DriverPack metadata from `HPClientDriverPackCatalog.cab` and WinPE pack metadata from HP's WinPE page.

- Script: `Scripts/Update-HPCatalog.ps1`
- Schema: `Schemas/HPCatalog.schema.json`
- Outputs:
  - `Cache/DriverPack/HP/DriverPack_HP.json`
  - `Cache/DriverPack/HP/DriverPack_HP.xml`
  - `Cache/DriverPack/HP/README.md`
  - `Cache/WinPE/HP/WinPE_HP.json`
  - `Cache/WinPE/HP/WinPE_HP.xml`
  - `Cache/WinPE/HP/README.md`

Run:

```powershell
pwsh -NoProfile -File ./Scripts/Update-HPCatalog.ps1
```
