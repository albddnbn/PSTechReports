function Scan-ForAppOrFilePath {
    <#
    .SYNOPSIS
        Scan a group of computers for a specified file/folder or application, and output the results to a .csv and .xlsx report.

    .DESCRIPTION
        The script searches application DisplayNames when the -type 'app' argument is used, and searches for files/folders when the -type 'path' argument is used.
        Creates a .csv/.xlsx report with findings, or outputs to gridview depending on $OutputFile parameter value.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER Item
        The item to search for.
        If the -SearchType 'app' argument is used, this should be the application's DisplayName.
        If the -SearchType 'path' argument is used, this should be the path to search for, Ex: C:\users\public\test.txt.

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER SearchType
        The type of search to perform.
        This can be either 'app' or 'path'.
        If 'app' is specified, the script will search for the specified application in the registry.
        If 'path' is specified, the script will search for the specified file/folder path on the target's filesystem.

    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .EXAMPLE
        Scan specified computers for Microsoft Teams application and output findings to 'teams.csv/.xlsx' files.
        Scan-ForAppOrFilePath ComputerName 't-client-01' -Item 'Microsoft Teams' -outputfile 'teams' -App

    .EXAMPLE
        Scan specified computers for presence of C:\users\public\test.txt file and output findings to 'testfile.csv/.xlsx' files.
        Scan-ForAppOrFilePath ComputerName 't-client-' -Item 'C:\users\public\test.txt' -outputfile 'testfile' -Path
    #>

    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [Parameter(Mandatory = $true)]
        [String]$Item,
        [Parameter(Mandatory = $true)]
        [String]$Outputfile,
        [switch]$SendPings,
        [switch]$App,
        [switch]$Path
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName


    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    if ($Path) {

        $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            $obj = [PSCustomObject]@{
                Name           = $env:COMPUTERNAME
                Path           = $using:item
                PathPresent    = $false
                PathType       = $null
                LastWriteTime  = $null
                CreationTime   = $null
                LastAccessTime = $null
                Attributes     = $null
            }
            $GetSpecifiedItem = Get-Item -Path "$using:item" -ErrorAction SilentlyContinue
            if ($GetSpecifiedItem.Exists) {
                $details = $GetSpecifiedItem | Select-Object FullName, *Time, Attributes, Length
                $obj.PathPresent = $true
                if ($GetSpecifiedItem.PSIsContainer) {
                    $obj.PathType = 'Folder'
                }
                else {
                    $obj.PathType = 'File'
                }
                $obj.LastWriteTime = $details.LastWriteTime
                $obj.CreationTime = $details.CreationTime
                $obj.LastAccessTime = $details.LastAccessTime
                $obj.Attributes = $details.Attributes
            }
            else {
                $obj.PathPresent = "Filepath not found"
            }
            $obj
        } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername

    }
    ## Application search
    elseif ($App) {

        $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
            # $app_matches = [System.Collections.ArrayList]::new()
            # Define the registry paths for uninstall information
            $registryPaths = @(
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            $obj = $null
            # Loop through each registry path and retrieve the list of subkeys
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
                    if ($displayName -like "*$using:Item*") {
                        $uninstallString = (Get-ItemProperty -Path $keyPath -Name "UninstallString" -ErrorAction SilentlyContinue).UninstallString
                        $version = (Get-ItemProperty -Path $keyPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
                        $publisher = (Get-ItemProperty -Path $keyPath -Name "Publisher" -ErrorAction SilentlyContinue).Publisher
                        $installLocation = (Get-ItemProperty -Path $keyPath -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
                        # $productcode = (Get-ItemProperty -Path $keyPath -Name "productcode" -ErrorAction SilentlyContinue).productcode
                        $installdate = (Get-ItemProperty -Path $keyPath -Name "installdate" -ErrorAction SilentlyContinue).installdate

                        $obj = [PSCustomObject]@{
                            ComputerName    = $env:COMPUTERNAME
                            AppName         = $displayName
                            AppVersion      = $version
                            InstallDate     = $installdate
                            InstallLocation = $installLocation
                            Publisher       = $publisher
                            UninstallString = $uninstallString
                        }
                        $obj
                    }
                }
            }
            # if ($null -eq $obj) {
            #     $obj = [PSCustomObject]@{
            #         ComputerName    = $single_computer
            #         AppName         = "No matching apps found for $using:Item"
            #         AppVersion      = $null
            #         InstallDate     = $null
            #         InstallLocation = $null
            #         Publisher       = $null
            #         UninstallString = "No matching apps found"
            #     }
            #     $obj
            # }
        } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No search type specified, exiting."
        return
    }

    ## Tries to collect hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results.count -ge 1) {
        $outputfile = Join-Path -Path $REPORT_DIRECTORY -ChildPath $outputfile

        $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
        "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
        $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel

            Import-CSV "$outputfile.csv" | Export-Excel -Path "$outputfile.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow

        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
        }

        Invoke-item "$outputfile.csv"
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."
    }

    return $results
}