# vybe-test: powershell/select_string_cmdlet/select_string_no_matches_returns_null
# When no matches are found in input objects, Select-String evaluates to $null
$result = @("alpha", "beta", "gamma") | Select-String -Pattern "zeta"

if ($null -ne $result) {
    Write-Host "FAIL: expected `$null when pattern does not match, got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
