function Scan-SoftwareInventory {
    <#
    .SYNOPSIS
        Scans a group of computers for installed applications and exports results to .csv/.xlsx - one per computer.

    .DESCRIPTION
        Scan-SoftwareInventory can handle a single string hostname as a target, a single string filepath to hostname list, or an array/arraylist of hostnames.
        The function uses 'AppsToLookFor' parameter to allow user to specify certain applications to look for in a comma-separated list.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER Outputfile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER AppsToLookFor
        Comma-separated list.
        Optional parameter to specify a list of applications/strings to look for. If not specified, all applications are scanned.

    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .EXAMPLE
        Scan-SoftwareInventory -ComputerName "t-client-28" -Title "tclient28-software"
    #>

    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [Parameter(
            Mandatory = $true)]
        [string]$OutputFile,
        $AppsToLookFor,
        [switch]$SendPings
    )
    if ($AppsToLookFor) {
        $AppsToLookFor = $AppsToLookFor.split(",")
        if ($AppsToLookFor -isnot [array]) {
            $AppsToLookFor = @($AppsToLookFor)
        }
    }

    $ComputerName = GetTargets -TargetComputer $ComputerName


    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    $results = invoke-command -computername $ComputerName -scriptblock {

        $targetapps = ($using:AppsToLookFor)
        $registryPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        foreach ($path in $registryPaths) {
            $uninstallKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            # Skip if the registry path doesn't exist
            if (-not $uninstallKeys) {
                continue
            }
            # Loop through each uninstall key and display the properties
            foreach ($key in $uninstallKeys) {
                $keyPath = Join-Path -Path $path -ChildPath $key.PSChildName
                $displayName = (Get-ItemProperty -Path $keyPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
                $uninstallString = (Get-ItemProperty -Path $keyPath -Name "UninstallString" -ErrorAction SilentlyContinue).UninstallString
                $version = (Get-ItemProperty -Path $keyPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
                $publisher = (Get-ItemProperty -Path $keyPath -Name "Publisher" -ErrorAction SilentlyContinue).Publisher
                $installLocation = (Get-ItemProperty -Path $keyPath -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
                $productcode = (Get-ItemProperty -Path $keyPath -Name "productcode" -ErrorAction SilentlyContinue).productcode
                $installdate = (Get-ItemProperty -Path $keyPath -Name "installdate" -ErrorAction SilentlyContinue).installdate
                $application_size = $null ## define as null for each loopthru


                if (($displayname -ne '') -and ($null -ne $displayname)) {
                    # if a target app list was provided, cycle through it and see if we're dealing with an app installation that is being searched for.
                    if ($targetapps) {
                        $matched_app = $false

                        $targetapps | ForEach-Object {
                            if ($displayname -like "*$_*") {
                                $matched_app = $true
                            }
                        }
                        ## If a search list was provided and there was no match, skip this app listing and move on to next
                        if (-not $matched_app) {
                            continue
                        }
                    }

                    ## Attempt to get approx 'size' of install location folder:
                    if ($installlocation) {
                        $application_size = (Get-ChildItem -Path "$installLocation" -Recurse -ErrorAction SilentlyContinue | MEasure-Object -Property Length -Sum -ErrorAction SilentlyContinue).sum / 1GB
                        $application_size = [Math]::Round($application_size, 2)
                        $application_size = "$application_size GB"
                    }

                    $obj = [pscustomobject]@{
                        DisplayName     = $displayName
                        UninstallString = $uninstallString
                        Version         = $version
                        Publisher       = $publisher
                        InstallLocation = $installLocation
                        ProductCode     = $productcode
                        InstallDate     = $installdate
                        ApplicationSize = $application_size
                    }
                    $obj
                }
            }
        }
    } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername

    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## Outputs results
    if ($results.count -ge 1) {
        $unique_hostnames = $($results.pscomputername) | Select-Object -Unique

        if ($errored_machines.count -ge 1) {
            Write-Host "These machines errored out during Invoke-Command." -ForegroundColor Red
            $errored_machines
        }

        ForEach ($single_computer_name in $unique_hostnames) {
            # get that computers apps
            $apps = $results | where-object { $_.pscomputername -eq $single_computer_name }
            # create the full filepaths
            $output_filepath = "$outputfile-$single_computer_name"
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Exporting files for $single_computername to $output_filepath."

            $apps | Export-Csv -Path "$outputfile-$single_computer_name.csv" -NoTypeInformation
            Import-CSV "$outputfile-$single_computer_name.csv" | Export-Excel -Path "$outputfile-$single_computer_name.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow


        }
        ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should exist
        try {
            Invoke-item "$($outputfile | split-path -Parent)"
        }
        catch {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder, attempting to open first .csv in list." -Foregroundcolor Yellow
            Invoke-item "$outputfile-$($unique_hostnames | Select-Object -first 1).csv"
        }
    }

    return $results
}