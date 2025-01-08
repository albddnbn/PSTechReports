function Get-InventoryDetails {
    <#
    .SYNOPSIS
        Targets computers and collects details relating to inventory like asset tag, serial number, and monitor details.

    .DESCRIPTION
        This has mainly been tested with Dell equipment - computers and monitors.
        Targets computers and collects details relating to inventory like asset tag, serial number, and monitor details.
        Creates a .csv/.xlsx report with findings, or outputs to gridview depending on $OutputFile parameter value.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and whether the Teams/Zoom processes are running.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Get-InventoryDetails -ComputerName "t-client-01" -Outputfile "tclient-01-details"

    .EXAMPLE
        Get-InventoryDetails -ComputerName "t-client-" -Outputfile "tclient-details"
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile,
        [switch]$SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName



    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    $gridview_title = "Inventory"

    $results = Invoke-Command -ComputerName $ComputerName -scriptblock {
        $pc_asset_tag = Get-Ciminstance -class win32_systemenclosure | Select-Object -exp smbiosassettag
        $pc_model = Get-Ciminstance -class win32_computersystem | Select-Object -exp model
        $pc_serial = Get-Ciminstance -class Win32_SystemEnclosure | Select-Object -exp serialnumber
        $pc_manufacturer = Get-Ciminstance -class Win32_ComputerSystem | Select-Object -exp manufacturer
        $monitors = Get-CimInstance WmiMonitorId -Namespace root\wmi | Select-Object SerialNumberID, ManufacturerName, UserFriendlyName
        $monitors | ForEach-Object {
            # $_.serialnumberid = [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -notmatch 0)
            #
            $_.UserFriendlyName = [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName)
            if ($_.UserFriendlyName -like "*P19*") {
                $_.serialnumberid = $(([System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -notmatch 0)).Trim())
            }
            else {
                ## from copilot: his will replace any character that is not in the range from hex 20 (space) to hex 7E (tilde), which includes all printable ASCII characters, with nothing.
                $_.serialnumberid = ($([System.Text.Encoding]::ASCII.GetString($_.SerialNumberID ).Trim()) -replace '[^\x20-\x7E]', '')
            }

            $_.ManufacturerName = [System.Text.Encoding]::ASCII.GetString($_.ManufacturerName)
        }

        $obj = [pscustomobject]@{

            computer_asset        = $pc_asset_tag
            computer_location     = $(($env:COMPUTERNAME -split '-')[1]) ## at least make an attempt to get location.
            computer_model        = $pc_model
            computer_serial       = $pc_serial
            computer_manufacturer = $pc_manufacturer
            monitor_serials       = $(($monitors.serialnumberid) -join ',')
            monitor_manufacturers = $(($monitors.ManufacturerName) -join ',')
            monitor_models        = $(($monitors.UserFriendlyName) -join ',')
            inventoried           = $true
        }
        $obj
    } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty PSShowComputerName, RunspaceId

    $not_inventoried = $ComputerName | Where-Object { $_ -notin $results.pscomputername }
    $not_inventoried += $RemoteError.CategoryInfo.TargetName | Where-Object { $_ -notin $not_inventoried }

    ## This section will attempt to output a CSV and XLSX report if anything other than 'n' was used for $Outputfile.
    ## If $Outputfile = 'n', results will be displayed in a gridview, with title set to $gridview_title.
    if ($results.count -ge 1) {

        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -title $gridview_title
        }
        else {

            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            ## Try ImportExcel
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel

                Import-CSV "$outputfile.csv" | Export-Excel -Path "$outputfile.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }

            Invoke-item "$outputfile.csv"
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Sample-Function." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}