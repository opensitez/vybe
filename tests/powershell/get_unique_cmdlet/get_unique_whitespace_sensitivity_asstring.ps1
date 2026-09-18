# vybe-test: powershell/get_unique_cmdlet/get_unique_whitespace_sensitivity_asstring
# -AsString preserves whitespace differences between strings
$items = @("  abc  ", "  abc  ", "abc")
$unique = @($items | Get-Unique -AsString)

if ($unique.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne "  abc  " -or $unique[1] -ne "abc") {
    Write-Host "FAIL: whitespace sensitive values mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
