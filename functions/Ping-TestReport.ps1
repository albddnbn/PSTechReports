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

    if ($results.count -ge 1) {

        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title $gridview_title
        }
        else {
            $outputfile = Join-Path -Path $REPORT_DIRECTORY -ChildPath $outputfile

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