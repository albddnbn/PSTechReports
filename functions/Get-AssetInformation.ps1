function Get-AssetInformation {
    <#
    .SYNOPSIS
        Attempts to use Dell Command Configure to get asset tag from BIOS, and other relevant information.

    .DESCRIPTION
        Collects: Computer model, BIOS version, BIOS release date, asset tag/serial number, and connected monitor information.
        Function will attempt to create .csv/.xlsx report, and return an arraylist with the results.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .EXAMPLE
        Get-AssetInformation -ComputerName s-c127-01 -Outputfile C127-01-AssetInfo

    .NOTES
        Issue: Errors being returned when gathering monitor details - it effects the 'errored_machines' hostname output.
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

    $gridview_title = "AssetInfo"

    ## Asset info scriptblock used to get local asset info from each target computer.
    $asset_info_scriptblock = {
        # computer model (ex: 'precision 3630 tower'), BIOS version, and BIOS release date
        $computer_model = get-ciminstance -class win32_computersystem | Select-Object -exp model
        $biosversion = get-ciminstance -class win32_bios | Select-Object -exp smbiosbiosversion
        $bioreleasedate = get-ciminstance -class win32_bios | Select-Object -exp releasedate
        try {
            $command_configure_exe = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Dell\Command Configure\x86_64" -Filter "cctk.exe" -File -ErrorAction Silentlycontinue
            # returns a string like: 'Asset=2001234'
            $asset_tag = &"$($command_configure_exe.fullname)" --asset
            $asset_tag = $asset_tag -replace 'Asset=', ''
        }
        catch {
            $asset_tag = Get-Ciminstance -class win32_systemenclosure | Select-Object -exp smbiosassettag
            # asus motherboard returned 'default string'
            if (($asset_tag.ToLower() -eq 'default string') -or (-not $asset_tag)) {
                $asset_tag = 'No asset tag set in BIOS'
            }
        }
        $computer_serial_num = get-ciminstance -class win32_bios | Select-Object -exp serialnumber
        # get monitor info and create a string from it (might be unnecessary, or a lengthy approach):
        $monitors = Get-CimInstance WmiMonitorId -Namespace root\wmi -ComputerName $ComputerName -ErrorAction SilentlyContinue
        if ($monitors) {
            $monitors = $monitors | Select-Object Active, ManufacturerName, UserFriendlyName, SerialNumberID, YearOfManufacture
            $monitor_string = ""
            $monitor_count = 0
            $monitors | ForEach-Object {
                $_.UserFriendlyName = [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName)
                $_.SerialNumberID = [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -notmatch 0)
                $_.ManufacturerName = [System.Text.Encoding]::ASCII.GetString($_.ManufacturerName)
                $manufacturername = $($_.ManufacturerName).trim()
                $monitor_string += "Maker: $manufacturername,Mod: $($_.UserFriendlyName),Ser: $($_.SerialNumberID),Yr: $($_.YearOfManufacture)"
                $monitor_count++
            }
        }
        else {
            $monitor_string = "No monitor information available."
            $monitor_count = 0
        }
        $obj = [PSCustomObject]@{
            model               = $computer_model
            biosversion         = $biosversion
            bioreleasedate      = $bioreleasedate
            asset_tag           = $asset_tag
            computer_serial_num = $computer_serial_num
            monitors            = $monitor_string
            NumMonitors         = $monitor_count
        }
        return $obj
    }

    $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock $asset_info_scriptblock -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername

    ## Tries to collect hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## If there were any results - output them to terminal and/or report files as necessary.
    if ($results.count -ge 1) {
        $results = $results | Sort-Object -property pscomputername
        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $outputfile = $outputfile | Select-Object -first 1
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            if ($errored_machines.count -ge 1) {
                "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"

                $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            }


            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel
                Import-CSV "$outputfile.csv" | Export-Excel -Path "$outputfile.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow

            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }

            Invoke-Item "$outputfile.csv"
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

    }
    return $results
}