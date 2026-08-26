# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_scriptblock_parameter
$sb = {
    param([ValidateNotNullOrEmpty()][string]$Arg)
    return "Arg:$Arg"
}
$res = & $sb -Arg "hello"
if ($res -ne "Arg:hello") {
    Write-Host "FAIL: ScriptBlock ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
