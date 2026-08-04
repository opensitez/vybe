# vybe-test: powershell/operators/join_operator_on_array
$parts = @("alpha", "beta", "gamma")
$result = $parts -join "-"
if ($result -ne "alpha-beta-gamma") {
    Write-Host "FAIL: '$result'"
    exit 1
}
# Empty string join
$nospace = $parts -join ""
if ($nospace -ne "alphabetagamma") {
    Write-Host "FAIL: no-sep '$nospace'"
    exit 1
}
Write-Host "PASS"
exit 0
