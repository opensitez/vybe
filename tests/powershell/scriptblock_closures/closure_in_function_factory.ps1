# vybe-test: powershell/scriptblock_closures/closure_in_function_factory
function New-Multiplier([int]$factor) {
    return { param($val) $val * $factor }.GetNewClosure()
}
$triple = New-Multiplier 3
$quad = New-Multiplier 4
if ( (&$triple 10) -ne 30 -or (&$quad 10) -ne 40 ) {
    Write-Host "FAIL: function closure factory expected 30 and 40"
    exit 1
}
Write-Host "PASS"
exit 0
