# ARI — Azure Resource Inventory Collector

A PowerShell wrapper around the open-source
[Azure Resource Inventory](https://github.com/microsoft/ARI) tool that generates an Excel
report and Draw.io diagram, then uploads both files to a secure customer path.

eGroup will provide you with your **CustomerName** and **SasToken** before you run this script.

## Run

Open a [Cloud Shell](https://shell.azure.com) instance and run:

```powershell
irm https://raw.githubusercontent.com/eGroupEnabling/ARI/main/ari.ps1 | iex
```

The script prompts for your `CustomerName` and `SasToken` if they are not supplied as parameters.

Parameterized run:

```powershell
irm https://raw.githubusercontent.com/eGroupEnabling/ARI/main/ari.ps1 | iex
# or, if running locally:
pwsh ./ari.ps1 -CustomerName <provided> -SasToken '<provided>'
```

Generate a report locally without uploading:

```powershell
pwsh ./ari.ps1 -CustomerName <provided> -SkipUpload
```

## Requirements
- PowerShell 7 or later (`pwsh`)
- An authenticated Azure session (Cloud Shell satisfies this automatically)
- Permission to install PowerShell modules without admin rights
- The `CustomerName` and `SasToken` values provided by eGroup

## Verification

1. Confirm a new report file is generated (`*_Report_*.xlsx`).
2. Confirm a new diagram file is generated (`*_Diagram_*.xml`).
3. Confirm both uploads complete without HTTP errors.
4. Local output paths are shown at the end of the run.

## Notes

- Windows output path: `C:\AzureResourceInventory`
- Linux and Cloud Shell output path: `$HOME/AzureResourceInventory`
- The script clears the output directory before each run to avoid stale files
- Uploads retry up to 3 times on timeout
- Upload failures report the HTTP status and service response — share the full output with eGroup for troubleshooting
