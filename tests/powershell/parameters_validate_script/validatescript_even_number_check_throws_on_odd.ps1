# vybe-test: powershell/parameters_validate_script/validatescript_even_number_check_throws_on_odd
function Set-EvenNumber2 {
    param([ValidateScript({ $_ % 2 -eq 0 })][int]$Number)
    return $Number
}
$caught = $false
try {
    $x = Set-EvenNumber2 -Number 43
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when ValidateScript returns false"
    exit 1
}
Write-Host "PASS"
exit 0
