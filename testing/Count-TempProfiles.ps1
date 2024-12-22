
function Count-TempProfiles {
    <#
    .SYNOPSIS
        Assumes temporary user folder suffix is: "*.$env:USERDOMAIN*" unless otherwise specified.
        Counts the number of temporary user folders/profiles on each target computer.
        Outputs a gridview, excel spreadsheet, and/or csv file.
    
    .DESCRIPTION
        This function can be useful to find:
        1. Specific user accounts or profiles that are having issues on the network and need assistance.
        2. Specific computers on the network that are having issues syncing with domain / network shares.
        3. Specific files that cause issues with redirected folders and roaming user profiles.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER TempFolderSuffix
        Suffix for temporary user folders. Default is "*.$env:USERDOMAIN*"
        Ex: C:\Users\Tsmith28.LLDC.000

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects.
        Multiple objects can be output for one computer - each object will pertain to each user that has a temporary folder on that computer.
        Ex: One object (row in output table/grid), could be for user: Tsmith28 with folder count 2.
        This would signify that they have 2 temp folders on that computer in addition to any regular user folder.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Count-TempProfiles -TargetComputer 't-client-'
        Count-TempProfiles -TargetComputer 't-client-01,t-client-02,t-client-03' -OutputFile 'n'
        Count-TempProfiles -TargetComputer 'C:\users\public\computers.txt' -OutputFile 'A220'

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile = '',
        [string]$TempFolderSuffix = "$env:USERDOMAIN"
    )
    ## allow user to press enter for default suffix.
    if ($TempFolderSuffix -eq '') {
        $TempFolderSuffix = "$env:USERDOMAIN"
    }
    Write-Host "Temporary folder suffix set to: " -NoNewline
    Write-Host "$TempFolderSuffix" -ForegroundColor Yellow
    ## 4. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## 2. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "TempProfiles"
    if ($Outputfile.tolower() -eq 'n') {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected 'N' input for outputfile, skipping creation of outputfile."
    }
    else {
        if ($Outputfile.toLower() -eq '') {
            $REPORT_DIRECTORY = "$str_title_var"
        }
        else {
            $REPORT_DIRECTORY = $outputfile            
        }
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory (Get-Location).Path -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ForEach ($single_computer in $ComputerName) {
        if ($single_computer) {
            ## 2. Check for connectivity
            if ([System.IO.Directory]::Exists("\\$single_computer\c$\Users")) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is online." -Foregroundcolor Green

                # get all temp profiles on computer / count / create object
                $temp_profile_folders = Get-Childitem -Path "\\$single_computer\c$\Users" -Filter "*.$env:USERDOMAIN*" -Directory -ErrorAction SilentlyContinue
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

    if ($results) {
        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -title $str_title_var
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            ## Try ImportExcel
            try {

                Import-Module ImportExcel

                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
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
    # read-host "Press enter to return results."
    return $results
}