function Copy-RemoteFiles {
    <#
    .SYNOPSIS
        Recursively grabs target files or folders from remote computer(s) and copies them to specified directory on local computer.

    .DESCRIPTION
        TargetPath specifies the target file(s) or folder(s) to target on remote machines.
        TargetPath can be supplied as a single absolute path, comma-separated list, or array.
        OutputPath specifies the directory to store the retrieved files.
        Creates a subfolder for each target computer to store it's retrieved files.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFolder
        Path to folder to store retrieved files. Ex: 'C:\users\abuddenb\Desktop\grabbed-files'

    .PARAMETER TargetPath
        Path to file(s)/folder(s) to be grabbed from remote machines. Ex: 'C:\users\abuddenb\Desktop\test.txt'

    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .EXAMPLE
        Copy the C:\test.txt file from target computers to C:\testgrab2 folder on local computer.
        Each remote computer will have it's own local subfolder in C:\testgrab2.
        Copy-RemoteFiles -ComputerName test-client -TargetPath C:\test.txt -OutputFolder C:\testgrab2 -SendPings

    .EXAMPLE
        Copy the C:\test folder from target computers to C:\testgrab2 folder on local computer.
        Each remote computer will have it's own local subfolder in C:\testgrab2.
        Copy-RemoteFiles -ComputerName test-client -TargetPath C:\test -OutputFolder C:\testgrab2 -SendPings

    #>
    param(        
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$TargetPath,
        [string]$OutputFolder,
        [switch]$SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    if (-not(Test-Path "$OutputFolder" -erroraction SilentlyContinue)) {
        New-Item -ItemType Directory -Path "$OutputFolder" -ErrorAction SilentlyContinue | out-null
    }

    ForEach ($single_computer in $ComputerName) {
        if ($single_computer) {
            $target_network_path = $targetpath -replace 'C:', "\\$single_computer\c$"
            write-host "Checking $single_computer...$target_network_path"

            if (Test-Path "$target_network_path" -erroraction SilentlyContinue) {

                    
                # $target_session = New-PSSession $single_computer

                $target_filename = $targetpath | split-path -leaf


                Copy-Item -Path "$target_network_path" -Destination "$OutputFolder\$single_computer-$target_filename"  -Recurse
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $targetpath ($single_computer) to $OutputFolder\$single_computer-$target_filename  complete." -foregroundcolor green
                    
                # Remove-PSSession $target_session

            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $targetpath on $single_computer to $OutputFolder on local computer (remote item may not exist)." -foregroundcolor red
            }
        }
    }
    ## Open output folder, pause.
    if (Test-Path "$OutputFolder" -erroraction SilentlyContinue) {
        Invoke-item "$OutputFolder"
    }
}