foreach ($directory in @('utilities', 'functions')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
}

Export-ModuleMember -Function *-*