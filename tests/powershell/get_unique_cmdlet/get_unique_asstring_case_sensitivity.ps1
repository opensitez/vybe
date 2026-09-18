# vybe-test: powershell/get_unique_cmdlet/get_unique_asstring_case_sensitivity
# In PowerShell Get-Unique, the -AsString switch performs CASE-SENSITIVE string comparisons
$items = @("apple", "Apple", "apple")
$unique = @($items | Get-Unique -AsString)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: -AsString is case-sensitive, expected 3 distinct casing records, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne "apple" -or $unique[1] -ne "Apple" -or $unique[2] -ne "apple") {
    Write-Host "FAIL: case-sensitive output mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
