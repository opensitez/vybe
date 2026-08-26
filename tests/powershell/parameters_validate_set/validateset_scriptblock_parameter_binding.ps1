# vybe-test: powershell/parameters_validate_set/validateset_scriptblock_parameter_binding
$sb = {
    param([ValidateSet("A", "B")][string]$Choice)
    return "Choice:$Choice"
}
$res = & $sb -Choice "B"
if ($res -ne "Choice:B") {
    Write-Host "FAIL: ScriptBlock ValidateSet failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
