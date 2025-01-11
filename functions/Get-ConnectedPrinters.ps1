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
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

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

    $list_local_printers_block = {
        $obj = [PScustomObject]@{
            Username          = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            DefaultPrinter    = $null
            ConnectedPrinters = $null
        }

        # Script will only check for printers if a user is logged in.
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

    if ($results.count -ge 1) {
        $results = $results | Sort-Object -property pscomputername
        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $outputfile = Join-Path -Path $REPORT_DIRECTORY -ChildPath $outputfile

            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation -Force
            if ($errored_machines.count -ge 1) {
                "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"

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