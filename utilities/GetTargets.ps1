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
    elseif ($TargetComputer.gettype().basetype.name -eq 'Array') {
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