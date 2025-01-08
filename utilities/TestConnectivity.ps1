

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
        if ($single_computer) {
            if (Test-Connection $single_computer -Count $PingCount -Quiet) {
                Write-Host "$single_computer is online." -ForegroundColor Green
                $online_results.Add($single_computer) | Out-Null
            }
            else {
                Write-Host "$single_computer is offline." -ForegroundColor Red
            }
        }
    }

    return $online_results

}