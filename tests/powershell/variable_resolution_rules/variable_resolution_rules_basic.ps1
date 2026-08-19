# vybe-test: powershell/variable_resolution_rules/basic
$outer = 'global'

function Resolve-Name {
    $outer = 'inner'
    return $outer
}

$inside = Resolve-Name
if ($inside -ne 'inner') {
    Write-Host "FAIL: inner scope read failed, got '$inside'"
    exit 1
}

if ($outer -ne 'global') {
    Write-Host "FAIL: outer variable leaked from function, got '$outer'"
    exit 1
}

Write-Host 'PASS'
exit 0
