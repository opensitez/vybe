# vybe-test: powershell/psprovider_introspection_engine/psprovider_wildcard_name_matching
# Get-PSProvider matches provider names using wildcard patterns
$fProviders = @(Get-PSProvider "F*")

if ($fProviders.Count -lt 2) {
    Write-Host "FAIL: expected at least 2 providers matching 'F*', got $($fProviders.Count)"
    exit 1
}

$names = @($fProviders | ForEach-Object { $_.Name })
if ($names -notcontains "FileSystem" -or $names -notcontains "Function") {
    Write-Host "FAIL: expected 'FileSystem' and 'Function', got: @($($names -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
