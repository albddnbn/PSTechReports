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
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

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