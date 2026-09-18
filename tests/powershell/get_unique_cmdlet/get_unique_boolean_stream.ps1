# vybe-test: powershell/get_unique_cmdlet/get_unique_boolean_stream
# Consecutive boolean elements are deduplicated in sequence
$bools = @($true, $true, $false, $false, $true)
$unique = @($bools | Get-Unique)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: expected 3 boolean transitions, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne $true -or $unique[1] -ne $false -or $unique[2] -ne $true) {
    Write-Host "FAIL: boolean transitions mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
