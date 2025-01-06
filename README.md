<img src="PSTechReports.png" alt="Concept drawing of PSTechReports icon"/>

# PSTechReports

The reports module from: https://github.com/albddnbn/PSTechSupportMenu.

More information: https://github.com/albddnbn/PSTechReports/wiki

## Functions available include:

- Get-AssetInformation
- Get-ComputerDetails
- Get-ConnectedPrinters
- Get-CurrentUser
- Get-InstalledDotNetversions
- Get-IntuneHardwareIDs
- Get-InventoryDetails
- Get-TempProfiles
- Ping-TestReport
- Scan-ForAppOrFilePath
- Scan-SoftwareInventory
- Test-ConnectivityQuick

## How GetTargets.ps1 works:

The **ComputerName** parameter will take a few things as input.
1. Single computer name
2. Comma-separated list of computer names
3. A substring of a computer name

OR

4. A path to text file containing list of computer names

**A regex pattern is used to validate computer name input for all values except text file.**

If you specify a "single computer name": s-client-1




### Get-IntuneHardwareIDs:

This function uses the script: https://www.powershellgallery.com/packages/Get-WindowsAutopilotInfo/3.8 to collect hardware IDs.

It allows you to submit a hostname text file, comma-separated list of hostnames, or 'hostname substring' to collect the IDs.
