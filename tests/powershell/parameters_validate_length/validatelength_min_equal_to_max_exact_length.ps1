# vybe-test: powershell/parameters_validate_length/validatelength_min_equal_to_max_exact_length
function Set-Pin {
    param([ValidateLength(4, 4)][string]$Pin)
    return $Pin
}
$r1 = Set-Pin -Pin "1234"
$caught = $false
try {
    $x = Set-Pin -Pin "12345"
} catch {
    $caught = $true
}
if ($r1 -ne "1234" -or -not $caught) {
    Write-Host "FAIL: Exact length ValidateLength failed"
    exit 1
}
Write-Host "PASS"
exit 0
