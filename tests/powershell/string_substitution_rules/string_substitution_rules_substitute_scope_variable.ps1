# vybe-test: powershell/string_substitution_rules/substitute_scope_variable
$var = 'root'
function Get-NestedValue {
    $var = 'child'
    return "$var"
}
if ((Get-NestedValue) -ne 'child') {
    Write-Host 'FAIL: substitution should use function-local scope value'
    exit 1
}
if ($var -ne 'root') {
    Write-Host "FAIL: parent scope variable changed unexpectedly to '$var'"
    exit 1
}

Write-Host 'PASS'
exit 0
