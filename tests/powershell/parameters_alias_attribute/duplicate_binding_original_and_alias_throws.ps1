# vybe-test: powershell/parameters_alias_attribute/duplicate_binding_original_and_alias_throws
function Set-SingleParam {
    param([Alias("A")][string]$ParamA)
    return $ParamA
}
$caught = $false
try {
    $x = Set-SingleParam -ParamA "one" -A "two"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Binding both parameter name and alias must fail"
    exit 1
}
Write-Host "PASS"
exit 0
