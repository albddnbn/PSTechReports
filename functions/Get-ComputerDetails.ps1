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
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

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
    $errored_machines

    read-host "Press Enter to continue"

    if ($results.count -ge 1) {
        $results = $results | Sort-Object -property pscomputername
        if (($outputfile.tolower() -eq 'n') -or (-not $outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation

            if ($errored_machines.count -ge 1) {
                "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
                $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
                Invoke-Item "$outputfile-Errors.csv"
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