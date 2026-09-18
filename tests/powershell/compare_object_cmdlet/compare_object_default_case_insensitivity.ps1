# vybe-test: powershell/compare_object_cmdlet/compare_object_default_case_insensitivity
# String comparisons in Compare-Object are case-insensitive by default
$diff = Compare-Object @("DATABASE_URL", "PORT") @("database_url", "port")

if ($null -ne $diff) {
    Write-Host "FAIL: strings differing only in casing unexpectedly flagged as different: $diff"
    exit 1
}

Write-Host "PASS"
exit 0
