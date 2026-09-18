# vybe-test: powershell/comparison/notlike_operator
# The -notlike operator returns true when a string does not match the wildcard pattern
$target = "PowerShell 7.4.0"

$mismatch = $target -notlike "*Python*"
$match = $target -notlike "*PowerShell*"

if (-not $mismatch) {
    Write-Host "FAIL: expected -notlike to return true when pattern does not match"
    exit 1
}

if ($match) {
    Write-Host "FAIL: expected -notlike to return false when pattern matches"
    exit 1
}

Write-Host "PASS"
exit 0
