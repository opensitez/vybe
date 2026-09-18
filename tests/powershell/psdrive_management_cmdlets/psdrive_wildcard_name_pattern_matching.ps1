# vybe-test: powershell/psdrive_management_cmdlets/psdrive_wildcard_name_pattern_matching
# Get-PSDrive supports wildcard patterns in the -Name parameter
$matches = @(Get-PSDrive -Name "En*")

if ($matches.Count -lt 1) {
    Write-Host "FAIL: expected at least 1 match for pattern 'En*', got 0"
    exit 1
}

if ($matches[0].Name -ne "Env") {
    Write-Host "FAIL: expected match 'Env', got: '$($matches[0].Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
