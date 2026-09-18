# vybe-test: powershell/compare_object_cmdlet/compare_object_casesensitive_switch_parameter
# -CaseSensitive forces strict case checking on string elements
$diff = Compare-Object @("DATABASE_URL") @("database_url") -CaseSensitive

if ($null -eq $diff) {
    Write-Host "FAIL: -CaseSensitive failed to detect casing difference"
    exit 1
}

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 diff records with -CaseSensitive, got $($diff.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
