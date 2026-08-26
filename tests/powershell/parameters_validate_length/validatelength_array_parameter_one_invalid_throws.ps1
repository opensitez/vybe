# vybe-test: powershell/parameters_validate_length/validatelength_array_parameter_one_invalid_throws
function Set-Aliases2 {
    param([ValidateLength(2, 4)][string[]]$Aliases)
    return $Aliases
}
$caught = $false
try {
    $x = Set-Aliases2 -Aliases "abc", "toolongalias"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when one array item exceeds length"
    exit 1
}
Write-Host "PASS"
exit 0
