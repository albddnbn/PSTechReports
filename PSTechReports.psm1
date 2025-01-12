foreach ($directory in @('utilities', 'functions')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
}

do {
    $REPORT_DIRECTORY = Read-Host "Enter path to reports directory: "
} until (Test-Path $REPORT_DIRECTORY -PathType Container -ErrorAction SilentlyContinue)


$REPORT_DIRECTORY = (Get-Item -Path $REPORT_DIRECTORY).FullName


Export-ModuleMember -Variable REPORT_DIRECTORY
Export-ModuleMember -Function *-*