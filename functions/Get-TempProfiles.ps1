function Get-TempProfiles {
    <#
    .SYNOPSIS
        Generates reports showing number of temporary folders found for a user, on each computer.
    
    .DESCRIPTION
        This function can be useful to find:
        1. Specific user accounts or profiles that are having issues on the network and need assistance.
        2. Specific computers on the network that are having issues syncing with domain / network shares.
        3. Specific files that cause issues with redirected folders and roaming user profiles.
        The function uses $env:USERDOMAIN as the default suffix for temporary folders, but this can be changed with the $TempFolderSuffix parameter.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        Path to output report. Script will add a .csv/.xlsx automatically.
        If 'n' is supplied, or Outputfile is not present in command, script will output findings to a gridview.

    .PARAMETER TempFolderSuffix
        Suffix for temporary user folders. Default is "*.$env:USERDOMAIN*"
        Ex: C:\Users\Tsmith28.LLDC.000
    
    .PARAMETER SendPings
        Switch parameter - if used will conduct ping test for connectivity on target computers before performing operations.
        Offline computers will be filtered out.

    .EXAMPLE
        Get-TempProfiles -TargetComputer 't-client-'
        Get-TempProfiles -TargetComputer 't-client-01,t-client-02,t-client-03' -OutputFile 'n'
        Get-TempProfiles -TargetComputer 'C:\users\public\computers.txt' -OutputFile 'A220'
    #>
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile = '',
        [string]$TempFolderSuffix = "$env:USERDOMAIN",
        [switch]$SendPings
    )
    ## Script will use the Domain netbios name for suffix if not provided.
    if (-not $TempFolderSuffix) {
        $TempFolderSuffix = "$env:USERDOMAIN"
    }
    Write-Host "Temporary folder suffix set to: " -NoNewline
    Write-Host "$TempFolderSuffix" -ForegroundColor Yellow

    $ComputerName = GetTargets -TargetComputer $ComputerName
    if ($SendPings) {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    $gridview_title = "TempProfiles"

    $results = [system.collections.arraylist]::new()

    ForEach ($single_computer in $ComputerName) {
        if ($single_computer) {
            ## Make sure remote computer is accessible
            if ([System.IO.Directory]::Exists("\\$single_computer\c$\Users")) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is online." -Foregroundcolor Green

                # get all temp profiles on computer / count / create object
                $temp_profile_folders = Get-Childitem -Path "\\$single_computer\c$\Users" -Filter "*.$TempFolderSuffix*" -Directory -ErrorAction SilentlyContinue

                $temp_profile_folders
                ## create object with  user, computer name, folder count to object, add to arraylist
                ForEach ($single_folder in $temp_profile_folders) {

                    $foldername = $single_folder.name

                    $username = $foldername.split('.')[0]
                    ## if the user and computer combo are not in results - add with count of 1
                    if ($results | Where-Object { ($_.User -eq $username) -and ($_.Computer -eq $single_computer) }) {
                        $results | Where-Object { ($_.User -eq $username) -and ($_.Computer -eq $single_computer) } | ForEach-Object { $_.FolderCount++ }
                        Write-Host "Found existing entry for $username and $single_computer increased FolderCount by 1."
                    }
                    else {
                        $temp_profile = [pscustomobject]@{
                            User        = $username
                            Computer    = $single_computer
                            FolderCount = 1
                        }
                        $temp_profile
                        $results.Add($temp_profile) | Out-Null
                        Write-Host "Added new entry for $username and $single_computer."
                    }
                }
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is offline." -Foregroundcolor Red
                continue
            }               
        }
    }

    if ($results.count -ge 1) {
        $results = $results | sort -property pscomputername

        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -title $gridview_title
        }
        else {

            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            ## Try ImportExcel
            try {

                Import-Module ImportExcel
                Import-CSV "$outputfile.csv" | Export-Excel -Path "$outputfile.xlsx" -AutoSize -TitleBackgroundColor Blue -TableStyle Medium9 -BoldTopRow
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-CurrentUser." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    return $results
}