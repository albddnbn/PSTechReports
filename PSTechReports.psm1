Function GetTargets {
    <#
    .SYNOPSIS
        Queries Active Directory for computer names based on hostname, hostname substring, or comma-separated list of hostnames.

    .DESCRIPTION
        Instead of using ActiveDirectory module, this function uses an LDAP Query to search for hostnames matching input.

    .PARAMETER TargetComputer
        Can be a single hostname, path to text file, comma-separated list of hostnames, or hostname substring which will 'grab' all hostnames starting with supplied substring.
        Ex: 't-client-' will grab all hostnames starting with 't-client-' (t-client-01, t-client-02, t-client-03, etc.)

    .EXAMPLE
        GetTargets -TargetComputer "t-client-01"

    .EXAMPLE
        GetTargets -TargetComputer "t-client-01,t-client-02"

    .EXAMPLE
        GetTargets -TargetComputer "t-client-"

    .EXAMPLE
        GetTargets -TargetComputer "D:\computers.txt"
    #>
    param(
        $TargetComputer
    )

    if ($TargetComputer -in @('', '127.0.0.1', 'localhost')) {
        $TargetComputer = @('127.0.0.1')
    }
    elseif ($(Test-Path $Targetcomputer -erroraction SilentlyContinue) -and ($TargetComputer.count -eq 1)) {
        $TargetComputer = Get-Content $TargetComputer
    }
    elseif ($TargetComputer.gettype().name -eq 'Array') {
        $TargetComputer = $TargetComputer
    }

    else {
        if ($Targetcomputer -like "*,*") {
            $TargetComputer = $TargetComputer -split ','
        }
        else {
            $Targetcomputer = @($Targetcomputer)
        }

        $NewTargetComputer = [System.Collections.Arraylist]::new()
        foreach ($computer in $TargetComputer) {
            ## CREDITS FOR The code this was adapted from: https://intunedrivemapping.azurewebsites.net/DriveMapping
            if ([string]::IsNullOrEmpty($env:USERDNSDOMAIN) -and [string]::IsNullOrEmpty($searchRoot)) {
                Write-Error "LDAP query `$env:USERDNSDOMAIN is not available!"
                Write-Warning "You can override your AD Domain in the `$overrideUserDnsDomain variable"
            }
            else {

                # if no domain specified fallback to PowerShell environment variable
                if ([string]::IsNullOrEmpty($searchRoot)) {
                    $searchRoot = $env:USERDNSDOMAIN
                }

                ## Thank you https://github.com/Jreece321 for this snippet - it shortened 10 lines of code to the 3 that you see below.
                $matching_hostnames = (([adsisearcher]"(&(objectCategory=Computer)(name=$computer*))").findall()).properties
                $matching_hostnames = $matching_hostnames.name
                $NewTargetComputer += $matching_hostnames
            }
        }
        $TargetComputer = $NewTargetComputer
    }

    $TargetComputer = $TargetComputer | Where-object { $_ -ne $null } | Select-Object -Unique
    if ($null -eq $TargetComputer) {
        return
    }
    return $TargetComputer
}

function TestConnectivity {
    <#
    .SYNOPSIS
        Tests connectivity to a single computer or list of computers by using Test-Connection -Quiet.

    .DESCRIPTION
        Tests connectivity to a single computer or list of computers by using Test-Connection -Quiet.
        Does not create any report, just gives green or red output to terminal based on ping response(s).

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PingCount
        Number of pings sent to each target machine. Default is 1.

    .EXAMPLE
        Check all hostnames starting with t-client- for online/offline status.
        TestConnectivity -ComputerName "t-client-"
    #>

    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        $PingCount = 1
    )
    $online_results = [system.collections.arraylist]::new()

    ## Ping target machines $PingCount times and log result to terminal.
    ForEach ($single_computer in $ComputerName) {
        if (Test-Connection $single_computer -Count $PingCount -Quiet) {
            Write-Host "$single_computer is online." -ForegroundColor Green
            $online_results.Add($single_computer) | Out-Null
        }
        else {
            Write-Host "$single_computer is offline." -ForegroundColor Red
        }
    }

    return $online_results

}


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
        Switch parameter to conduct ping test for connectivity before attempting main purpose of function.

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
    if ($results) {
        $results = $results | Sort-Object -property pscomputername
        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $outputfile = $outputfile | Select-Object -first 1
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            if ($errored_machines) {
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

function Get-ComputerDetails {
    <#
    .SYNOPSIS
        Collects: Manufacturer, Model, Current User, Windows Build, BIOS Version, BIOS Release Date, and Total RAM from target machine(s).
        Creates a .csv/.xlsx report with findings, or outputs to gridview depending on $OutputFile parameter value.

    .DESCRIPTION
        Collects: Manufacturer, Model, Current User, Windows Build, BIOS Version, BIOS Release Date, and Total RAM from target machine(s).
        Outputs: A .csv and .xlsx report file if anything other than 'n' is supplied for the $OutputFile parameter.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, computer model, bios version/release date, last boot time, and other info.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Output details for a single hostname to "sa227-28-details.csv" and "sa227-28-details.xlsx" in the 'reports' directory.
        Get-ComputerDetails -ComputerName "t-client-28" -Outputfile "tclient-28-details"

    .EXAMPLE
        Output details for all hostnames starting with g-pc-0 to terminal.
        Get-ComputerDetails -ComputerName 'g-pc-0'
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

    $gridview_title = "PCdetails"

    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        # Gets active user, computer manufacturer, model, BIOS version & release date, Win Build number, total RAM, last boot time, and total system up time.
        # object returned to $results list
        $computersystem = Get-CimInstance -Class Win32_Computersystem
        $bios = Get-CimInstance -Class Win32_BIOS
        $operatingsystem = Get-CimInstance -Class Win32_OperatingSystem

        $lastboot = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
        $uptime = ((Get-Date) - $lastboot).ToString("dd\.hh\:mm\:ss")
        $obj = [PSCustomObject]@{
            Manufacturer    = $($computersystem.manufacturer)
            Model           = $($computersystem.model)
            CurrentUser     = $((get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username)
            WindowsBuild    = $($operatingsystem.buildnumber)
            BiosVersion     = $($bios.smbiosbiosversion)
            BiosReleaseDate = $($bios.releasedate)
            TotalRAM        = $((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb)
            LastBoot        = $lastboot
            SystemUptime    = $uptime
        }
        $obj
    } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername -ErrorAction SilentlyContinue

    ## Tries to collect hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {
        $results = $results | Sort-Object -property pscomputername
        if (($outputfile.tolower() -eq 'n') -or (-not $outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation

            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            if ($errored_machines) {
                $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            }
            ## Try ImportExcel
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel

                Import-CSV "$outputfile.csv" | Export-Excel -Path "$outputfile.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }

            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
            Invoke-item "$outputfile.csv"
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-ComputerDetails." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    return $results
}

function Get-ConnectedPrinters {
    <#
    .SYNOPSIS
        Checks the target computer, and returns the user that's logged in, and the printers that user has access to.
        Creates a .csv/.xlsx report with findings, or outputs to gridview depending on $OutputFile parameter value.

    .DESCRIPTION
        This function, unlike some others, only takes a single string DNS hostname of a target computer.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER FolderTitleSubstring
        If specified, the function will create a folder in the 'reports' directory with the specified substring in the title, appended to the $outputfile String (relates to the function title).

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and list of connected printers.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Get-ConnectedPrinters -ComputerName 't-client-07'
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

    $gridview_title = "Printers"

    ## Scriptblock - lists connected/default printers
    $list_local_printers_block = {
        $obj = [PScustomObject]@{
            Username          = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            DefaultPrinter    = $null
            ConnectedPrinters = $null
        }

        # Only need to check for connected printers if a user is logged in.
        if ($obj.Username) {
            # get connected printers:
            get-ciminstance -class win32_printer | Select-Object name, Default | ForEach-Object {
                if (($_.name -notin ('Microsoft Print to PDF', 'Fax')) -and ($_.name -notlike "*OneNote*")) {
                    if ($_.name -notlike "Send to*") {
                        $obj.ConnectedPrinters = "$($obj.ConnectedPrinters), $($_.name)"
                    }
                }
            }
        }
        $obj
    }

    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock $list_local_printers_block  -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername

    ## Tries to collect hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {
        $results = $results | Sort-Object -property pscomputername
        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            if ($errored_machines) {
                $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            }

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

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-ConnectedPrinters." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}

function Get-CurrentUser {
    <#
    .SYNOPSIS
        Gets user logged into target system(s).

    .DESCRIPTION
        Creates report with current user, computer model, and if Teams or Zoom are running.
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
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and whether the Teams/Zoom processes are running.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        1. Get users on all S-A231 computers:
        Get-CurrentUser -ComputerName "s-a231-"

    .EXAMPLE
        2. Get user on a single target computer:
        Get-CurrentUser -ComputerName "t-client-28"

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
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


    $gridview_title = "CurrentUsers"

    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        $obj = [PSCustomObject]@{
            Model        = (get-ciminstance -class win32_computersystem).model
            CurrentUser  = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            TeamsRunning = $(if (Get-PRocess -Name 'Teams' -ErrorAction SilentlyContinue) { $true } else { $false })
            ZoomRunning  = $(if (Get-PRocess -Name 'Zoom' -ErrorAction SilentlyContinue) { $true } else { $false })

        }
        $obj
    } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername
    ## Tries to collect hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {

        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -title $gridview_title
        }
        else {

            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            if ($errored_machines) {
                $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            }

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

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-CurrentUser." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    return $results
}

function Get-InstalledDotNetversions {
    <#
    .SYNOPSIS
        Gets a list of installed dotnet versions on target computers and returns results.
        Uses Powershell from: https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed#query-the-registry-using-powershell-older-framework-versions
        And https://stackoverflow.com/questions/3487265/powershell-script-to-return-versions-of-net-framework-on-a-machine
        To return list of installed dotnets.

    .DESCRIPTION
        Gets a list of installed dotnet versions on target computers and returns results.
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
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing the hostname and info on installed .net versions.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        1. Get dotnet versions on single computer, output results to terminal/gridview
        Get-InstalledDotNetVersions -ComputerName "t-client-01" -outputfile 'n'

    .EXAMPLE
        2. Get user on group of computers with hostnames starting with t-client-, output default filename reports
        Get-InstalledDotNetVersions -ComputerName "t-client-" -outputfile ''

    .NOTES
        Sources include:
        https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed#query-the-registry-using-powershell-older-framework-versions
        https://stackoverflow.com/questions/3487265/powershell-script-to-return-versions-of-net-framework-on-a-machine  
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

    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        # Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse | `
        #     Get-ItemProperty -Name version -EA 0 | Where-Object { $_.PSChildName -Match '^(Where-Object !S)\p{L}' } |`
        #     Select-Object PSChildName, version

        $Lookup = @{
            378389 = [version]'4.5'
            378675 = [version]'4.5.1'
            378758 = [version]'4.5.1'
            379893 = [version]'4.5.2'
            393295 = [version]'4.6'
            393297 = [version]'4.6'
            394254 = [version]'4.6.1'
            394271 = [version]'4.6.1'
            394802 = [version]'4.6.2'
            394806 = [version]'4.6.2'
            460798 = [version]'4.7'
            460805 = [version]'4.7'
            461308 = [version]'4.7.1'
            461310 = [version]'4.7.1'
            461808 = [version]'4.7.2'
            461814 = [version]'4.7.2'
            528040 = [version]'4.8'
            528049 = [version]'4.8'
        }
        
        # For One True framework (latest .NET 4x), change the Where-Object match 
        # to PSChildName -eq "Full":
        Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse |
        Get-ItemProperty -name Version, Release -EA 0 |
        Where-Object { $_.PSChildName -match '^(?!S)\p{L}' } |
        Select-Object @{name = ".NET Framework"; expression = { $_.PSChildName } }, 
        @{name = "Product"; expression = { $Lookup[$_.Release] } }, 
        Version, Release

        Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' |
        Where-Object { ($_.PSChildName -ne "v4") -and ($_.PSChildName -like 'v*') } |
        ForEach-Object {
            $name = $_.Version
            $sp = $_.SP
            $install = $_.Install
            if (-not $install) {
                Write-Host -Object "$($_.PSChildName)  $($name)"
            }
            elseif ($install -eq '1') {
                if (-not $sp) {
                    Write-Host -Object "$($_.PSChildName)  $($name)"
                }
                else {
                    Write-Host -Object "$($_.PSChildName)  $($name) SP$($sp)"
                }
            }
            if (-not $name) {
                $parentName = $_.PSChildName
                Get-ChildItem -LiteralPath $_.PSPath |
                Where-Object {
                    if ($_.Property -contains 'Version') { $name = $((Get-ItemProperty -Path "Registry::$_").Version) }
                    if ($name -and ($_.Property -contains 'SP')) { $sp = $((Get-ItemProperty -Path "Registry::$_").SP) }
                    if ($_.Property -contains 'Install') { $install = $((Get-ItemProperty -Path "Registry::$_").Install) }
                    if (-not $install) {
                        Write-Host -Object "  $($parentName)  $($name)"
                    }
                    elseif ($install -eq '1') {
                        if (-not $sp) {
                            Write-Host -Object "  $($_.PSChildName)  $($name)"
                        }
                        else {
                            Write-Host -Object "  $($_.PSChildName)  $($name) SP$($sp)"
                        }
                    }
                }
            }
        }
        

    } -ErrorVariable RemoteError | Select-Object * -ExcludeProperty RunspaceId, PSshowcomputername

    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {

        # ForEach ($single_result in $results) {
        #     $single_result
        # }


        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title "Installed .NET Versions"
        }
        else {

            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            if ($errored_machines) {
                $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            }

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

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-InstalledDotNetVersions." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}

Function Get-IntuneHardwareIDs {
    <#
    .SYNOPSIS
        Uses Get-WindowsAutoPilotInfo to generate a .csv containing hardware ID info for target device(s), which can then be imported into Intune.

    .DESCRIPTION
        If $ComputerName = '', function is run on local computer.
        Specify GroupTag using DeviceGroupTag parameter.
        Uses Get-WindowsAutopilotInfo from: https://github.com/MikePohatu/Get-WindowsAutoPilotInfo/blob/main/Get-WindowsAutoPilotInfo.ps1
        Get-WindowsAutopilotInfo.ps1 is in the supportfiles directory, so it doesn't have to be installed/downloaded from online.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER DeviceGroupTag
        Specifies the group tag that will be set in target devices' hardware ID info.
        DeviceGroupTag value is used with the -GroupTag parameter of Get-WindowsAutopilotInfo.

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        Outputs .csv file containing HWID information for target devices, to upload them into Intune.

    .EXAMPLE
        Get Intune Hardware IDs from all computers in room A227 on Stanton campus:
        Get-IntuneHardwareIDs -ComputerName "t-client-" -OutputFile "TClientIDs" -DeviceGroupTag 'Student Laptops'

    .EXAMPLE
        Get Intune Hardware ID of single target computer
        Get-IntuneHardwareIDs -ComputerName "t-client-01" -OutputFile "TClient01-ID"
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile,
        [string]$DeviceGroupTag,
        [switch]$SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName


    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    ## make sure there's a .csv
    if ($outputfile -notlike "*.csv") {
        $outputfile += ".csv"
    }


    ## Installs the script and then gets absolute path to execute it.
    $check_for_nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if ($null -eq $check_for_nuget) {
        # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: NuGet not found, installing now."
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    }
    Install-Script -Name 'Get-WindowsAutopilotInfo' -Force -requiredversion 3.8


    ## Define parameters to be used when executing Get-WindowsAutoPilotInfo
    $params = @{
        ComputerName = $ComputerName
        OutputFile   = "$outputfile"
        GroupTag     = $DeviceGroupTag
        Append       = $true
    }

    $script_path = Get-InstalledScript -Name Get-WindowsAutoPilotInfo | Select-Object -Exp InstalledLocation
    &"$script_path\Get-WindowsAutoPilotInfo.ps1" @params

    ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
    try {
        Invoke-item "$($outputfile | split-path -Parent)"
    }
    catch {
        Invoke-item "$outputfile"
    }

}

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
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

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
    if ($results) {

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

function Ping-TestReport {
    <#
    .SYNOPSIS
        Pings a group of computers a specified amount of times, and outputs the successes / total pings to a .csv and .xlsx report.

    .DESCRIPTION
        Pings a group of computers a specified amount of times, and outputs the successes / total pings to a .csv and .xlsx report.
        Creates a .csv/.xlsx report with findings, or outputs to gridview depending on $OutputFile parameter value.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PingCount
        Number of times to ping each computer.

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .EXAMPLE
        Ping-TestReport -ComputerName "g-client-" -PingCount 10 -Outputfile "GClientPings"

    .EXAMPLE
        Ping-TestReport -ComputerName "g-client-" -PingCount 2

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>

    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        $PingCount,
        [string]$Outputfile = ''
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    $am_pm = (Get-Date).ToString('tt')

    $gridview_title = "Pings-$Outputfile-$(Get-Date -Format 'hh-MM')$($am_pm)"

    ## Create arraylist to store results
    $results = [system.collections.arraylist]::new()

    $PingCount = [int]$PingCount


    ForEach ($single_computer in $ComputerName) {

        if ($single_computer) {

            ## check if network path exists first - that way we don't waste time pinging machine thats offlineWhere-Object
            if (-not ([System.IO.Directory]::Exists("\\$single_computer\c$"))) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is not online." -foregroundcolor red
                continue
            }

            ## Create object to store results of ping test on single machine
            $obj = [pscustomobject]@{
                Sourcecomputer       = $env:COMPUTERNAME
                ComputerHostName     = $single_computer
                TotalPings           = $pingcount
                Responses            = 0
                AvgResponseTime      = 0
                PacketLossPercentage = 0
            }
            Write-Host "Sending $pingcount pings to $single_computer..."
            ## Send $PINGCOUNT number of pings to target device, store results
            $send_pings = Test-Connection -ComputerName $single_computer -count $PingCount -ErrorAction SilentlyContinue
            ## Set number of responses from target machine
            $obj.responses = $send_pings.count
            ## Calculate average response time for successful responses
            $sum_of_response_times = $($send_pings | measure-object responsetime -sum)
            if ($obj.Responses -eq 0) {
                $obj.AvgResponseTime = 0
            }
            else {
                $obj.avgresponsetime = $sum_of_response_times.sum / $obj.responses
            }
            ## Calculate packet loss percentage - divide total pings by responses
            $total_drops = $obj.TotalPings - $obj.Responses
            $obj.PacketLossPercentage = ($total_drops / $($obj.TotalPings)) * 100

            $results.add($obj) | Out-Null
        }
    }

    if ($results) {

        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title $gridview_title
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
    }

    return $results
}

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
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        Scan-ForAppOrFilePath ComputerName 't-client-01' -SearchType 'app' -Item 'Microsoft Teams' -outputfile 'teams'
    #>

    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [Parameter(Mandatory = $true)]
        [String]$Item,
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

    if ($results) {
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
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

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
    if ($results) {
        $unique_hostnames = $($results.pscomputername) | Select-Object -Unique

        if ($errored_machines) {
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

function Test-ConnectivityQuick {
    <#
    .SYNOPSIS
        Tests connectivity to a single computer or list of computers by using Test-Connection -Quiet.

    .DESCRIPTION
        Works fairly quickly, but doesn't give you any information about the computer's name, IP, or latency - judges online/offline by the 1 ping.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PingCount
        Number of pings sent to each target machine. Default is 1.

    .EXAMPLE
        Check all hostnames starting with t-client- for online/offline status.
        TestConnectivityQuick -ComputerName "t-client-"
    #>

    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        $PingCount = 1
    )

    $PingCount = $PingCount

    if ($null -eq $ComputerName) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected pipeline for targetcomputer." -Foregroundcolor Yellow
    }
    else {
        ## Assigns localhost value
        if ($ComputerName -in @('', '127.0.0.1', 'localhost')) {
            $ComputerName = @('127.0.0.1')
        }
        ## If input is a file, gets content
        elseif ($(Test-Path $ComputerName -erroraction SilentlyContinue) -and ($ComputerName.count -eq 1)) {
            $ComputerName = Get-Content $ComputerName
        }
        ## A. Separates any comma-separated strings into an array, otherwise just creates array
        ## B. Then, cycles through the array to process each hostname/hostname substring using LDAP query
        else {
            ## A.
            if ($ComputerName -like "*,*") {
                $ComputerName = $ComputerName -split ','
            }
            else {
                $ComputerName = @($ComputerName)
            }

            ## B. LDAP query each ComputerName item, create new list / sets back to ComputerName when done.
            $NewTargetComputer = [System.Collections.Arraylist]::new()
            foreach ($computer in $ComputerName) {
                ## CREDITS FOR The code this was adapted from: https://intunedrivemapping.azurewebsites.net/DriveMapping
                if ([string]::IsNullOrEmpty($env:USERDNSDOMAIN) -and [string]::IsNullOrEmpty($searchRoot)) {
                    Write-Error "LDAP query `$env:USERDNSDOMAIN is not available!"
                    Write-Warning "You can override your AD Domain in the `$overrideUserDnsDomain variable"
                }
                else {

                    # if no domain specified fallback to PowerShell environment variable
                    if ([string]::IsNullOrEmpty($searchRoot)) {
                        $searchRoot = $env:USERDNSDOMAIN
                    }

                    $matching_hostnames = (([adsisearcher]"(&(objectCategory=Computer)(name=$computer*))").findall()).properties
                    $matching_hostnames = $matching_hostnames.name
                    $NewTargetComputer += $matching_hostnames
                }
            }
            $ComputerName = $NewTargetComputer
        }
        $ComputerName = $ComputerName | Where-object { $_ -ne $null } | Select-Object -Unique
        # Safety catch
        if ($null -eq $ComputerName) {
            return
        }
    }

    $results = [system.collections.arraylist]::new()

    ## Ping target machines $PingCount times and log result to terminal.
    ForEach ($single_computer in $ComputerName) {

        if ($single_computer) {
            $connection_result = Test-Connection $single_computer -count $PingCount -ErrorAction SilentlyContinue
            $ping_responses = $connection_result.count
            $ping_response_obj = [pscustomobject]@{
                ComputerName  = $single_computer
                Status        = ""
                PingResponses = $ping_responses
                NumberPings   = $PingCount
            }

            if ($connection_result) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is online [$ping_responses responses]" -foregroundcolor green
                $ping_response_obj.Status = 'online'
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: " -NoNewline
                Write-Host "$single_computer is not online." -foregroundcolor red
                $ping_response_obj.Status = 'offline'
            }

            $results.add($ping_response_obj) | Out-Null
        }

    }
    ## Open results in gridview since this is just supposed to be quick test for connectivity
    $results | out-gridview -Title "Results: $PingCount Pings"
}

function Count-TempProfiles {
    <#
    .SYNOPSIS
        Generates reports showing number of temporary folders found for a user, on each computer.
    
    .DESCRIPTION
        This function can be useful to find:
        1. Specific user accounts or profiles that are having issues on the network and need assistance.
        2. Specific computers on the network that are having issues syncing with domain / network shares.
        3. Specific files that cause issues with redirected folders and roaming user profiles.
        The function uses $env:USERDOMAIN as the default suffix for temporary folders, but this can be changed with the $TempFolderSuffix parameter.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER TempFolderSuffix
        Suffix for temporary user folders. Default is "*.$env:USERDOMAIN*"
        Ex: C:\Users\Tsmith28.LLDC.000

    .EXAMPLE
        Count-TempProfiles -TargetComputer 't-client-'
        Count-TempProfiles -TargetComputer 't-client-01,t-client-02,t-client-03' -OutputFile 'n'
        Count-TempProfiles -TargetComputer 'C:\users\public\computers.txt' -OutputFile 'A220'
    #>
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile = '',
        [string]$TempFolderSuffix = "$env:USERDOMAIN",
        [switch]$SendPings
    )
    ## Script will use the Domain netbios name for suffix if not provided.
    if (-not $TempFolderSuffix) {
        $TempFolderSuffix = "$env:USERDOMAIN"
    }
    Write-Host "Temporary folder suffix set to: " -NoNewline
    Write-Host "$TempFolderSuffix" -ForegroundColor Yellow

    $ComputerName = GetTargets -TargetComputer $ComputerName
    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    $gridview_title = "TempProfiles"

    $results = [system.collections.arraylist]::new()

    ForEach ($single_computer in $ComputerName) {
        if ($single_computer) {
            ## Make sure remote computer is accessible
            if ([System.IO.Directory]::Exists("\\$single_computer\c$\Users")) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is online." -Foregroundcolor Green

                # get all temp profiles on computer / count / create object
                $temp_profile_folders = Get-Childitem -Path "\\$single_computer\c$\Users" -Filter "*.$TempFolderSuffix*" -Directory -ErrorAction SilentlyContinue

                $temp_profile_folders
                ## create object with  user, computer name, folder count to object, add to arraylist
                ForEach ($single_folder in $temp_profile_folders) {

                    $foldername = $single_folder.name

                    $username = $foldername.split('.')[0]
                    ## if the user and computer combo are not in results - add with count of 1
                    if ($results | Where-Object { ($_.User -eq $username) -and ($_.Computer -eq $single_computer) }) {
                        $results | Where-Object { ($_.User -eq $username) -and ($_.Computer -eq $single_computer) } | ForEach-Object { $_.FolderCount++ }
                        Write-Host "Found existing entry for $username and $single_computer increased FolderCount by 1."
                    }
                    else {
                        $temp_profile = [pscustomobject]@{
                            User        = $username
                            Computer    = $single_computer
                            FolderCount = 1
                        }
                        $temp_profile
                        $results.Add($temp_profile) | Out-Null
                        Write-Host "Added new entry for $username and $single_computer."
                    }
                }
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is offline." -Foregroundcolor Red
                continue
            }               
        }
    }

    if ($results) {
        $results = $results | sort -property pscomputername

        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -title $gridview_title
        }
        else {

            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            ## Try ImportExcel
            try {

                Import-Module ImportExcel
                Import-CSV "$outputfile.csv" | Export-Excel -Path "$outputfile.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-CurrentUser." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    return $results
}

Export-ModuleMember -Function *-*