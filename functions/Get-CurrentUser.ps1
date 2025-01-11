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
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and whether the Teams/Zoom processes are running.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        1. Get users on all S-A231 computers:
        Get-CurrentUser -ComputerName "s-a231-"

    .EXAMPLE
        2. Get user on a single target computer:
        Get-CurrentUser -ComputerName "t-client-28"
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

    if ($results.count -ge 1) {

        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -title $gridview_title
        }
        else {
            $outputfile = Join-Path -Path $REPORT_DIRECTORY -ChildPath $outputfile

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