# vybe-test: powershell/variable_lifetime_and_scoping/basic
function Scoped-Create {
    $localValue = 'block-scope'
    return $localValue
}

$returned = Scoped-Create
if ($returned -ne 'block-scope') {
    Write-Host "FAIL: function return wrong, got '$returned'"
    exit 1
}

if (Get-Variable -Name localValue -Scope Local -ErrorAction SilentlyContinue) {
    Write-Host 'FAIL: function-local variable leaked into caller scope'
    exit 1
}

Write-Host 'PASS'
exit 0
