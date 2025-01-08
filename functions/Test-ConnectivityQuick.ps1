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