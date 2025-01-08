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
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

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

    if ($results.count -ge 1) {

        # ForEach ($single_result in $results) {
        #     $single_result
        # }


        $results = $results | Sort-Object -property pscomputername

        if (($outputfile.tolower() -eq 'n') -or (-not $Outputfile)) {
            $results | out-gridview -Title "Installed .NET Versions"
        }
        else {

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

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-InstalledDotNetVersions." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}