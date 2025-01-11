foreach ($directory in @('utilities', 'functions')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
}

$REPORT_DIRECTORY = ""

while (-not (Test-Path $REPORT_DIRECTORY -PathType Container -ErrorAction SilentlyContinue)) {
    Write-Host "Directory does not exist. Please try again."
    $REPORT_DIRECTORY = Read-Host "Set Reports Directory: "
}

$REPORT_DIRECTORY = (Get-Item -Path $REPORT_DIRECTORY).FullName


Export-ModuleMember -Variable REPORT_DIRECTORY
Export-ModuleMember -Function *-*