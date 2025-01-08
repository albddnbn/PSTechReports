function Send-Files {
    <#
    .SYNOPSIS
        Sends a target file/folder from local computer to target path on remote computers.

    .DESCRIPTION
        You can enter both paths as if they're on local filesystem, the script should cut out any drive letters and insert the \\hostname\c$ for UNC path. The script only works for C drive on target computers right now.

    .PARAMETER SourcePath
        The path of the file/folder you want to send to target computers. 
        ex: C:\users\public\desktop\test.txt, 
        ex: \\networkshare\folder\test.txt

    .PARAMETER DestinationPath
        The path on the target computer where you want to send the file/folder. 
        The script will cut off any preceding drive letters and insert \\hostname\c$ - so destination paths should be on C drive of target computers.
        ex: C:\users\public\desktop\
        The path should be a destination folder, script should create directories/subdirectories as required.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.
        
    .EXAMPLE
        Send C:\test.txt file to target computers' C:\ drive (C:\test.txt)
        send-files -ComputerName test-client -sourcepath C:\test.txt -destinationpath C:\ -SendPings

    .EXAMPLE
        Send C:\test folder to target computers' C:\received folder (C:\received\test)
        send-files -ComputerName test-client -sourcepath C:\test -destinationpath C:\received -SendPings
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [ValidateScript({
                if (Test-Path $_ -ErrorAction SilentlyContinue) {
                    return $true
                }
                else {
                    Write-Error "SourcePath does not exist. Please enter a valid path."
                    return $false                
                }
            })]
        [string]$sourcepath,
        [string]$destinationpath,
        [switch]$SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    # create full output filepath for log file (shows confirmation that files were copied to each remote computer)
    $output_filepath = "SendFiles-$(get-date -format 'yyyy-MM-dd').txt"

    $informational_string = ""

    ForEach ($single_computer in $ComputerName) {
        if ($single_computer) {
            $file_copied = $false
            if ([System.IO.Directory]::Exists("\\$single_computer\c$")) {
                Invoke-Command -ComputerName $single_computer -Scriptblock {

                    $destination_folder = $using:destinationpath
                    if (-not (Test-Path $destination_folder -ErrorAction SilentlyContinue)) {
                        Write-Host "Creating $destination_folder directory on $env:COMPUTERNAME."
                        New-Item -Path $destination_folder -ItemType Directory -Force | Out-Null
                    }
                }

                $target_session = New-PSSession $single_computer

                Copy-Item -Path "$sourcepath" -Destination "$destinationpath" -ToSession $target_session -Recurse
                $informational_string += "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $sourcepath to $destinationpath ($single_computer) complete.`n"

                $file_copied = $true


                Remove-PSSession $target_session
                            
            }
            
            if (-not $file_copied) {
                $informational_string += "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $sourcepath to $destinationpath on $single_computer.`n"
            }
            
        }
    }
        
    ## Append text to file here:
    $informational_string | Out-File -FilePath $output_filepath -Append
    # $ComputerName | Out-File -FilePath $output_filepath -Append
    "`nThe Scan-ForApporFilepath function can be used to verify file/folders' existence on target computers." | Out-File -FilePath $output_filepath -Append -Force
        
    Invoke-Item "$output_filepath"        
}