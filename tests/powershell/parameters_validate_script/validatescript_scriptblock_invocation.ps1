# vybe-test: powershell/parameters_validate_script/validatescript_scriptblock_invocation
$sb = {
    param([ValidateScript({ $_ -match "^test_" })][string]$Name)
    return "Matched:$Name"
}
$res = & $sb -Name "test_item"
if ($res -ne "Matched:test_item") {
    Write-Host "FAIL: ScriptBlock ValidateScript failed"
    exit 1
}
Write-Host "PASS"
exit 0
