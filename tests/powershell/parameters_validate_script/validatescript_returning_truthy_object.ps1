# vybe-test: powershell/parameters_validate_script/validatescript_returning_truthy_object
function Check-TruthyObj {
    param([ValidateScript({ "any non-empty string is truthy" })][string]$Val)
    return "OK:$Val"
}
$res = Check-TruthyObj -Val "hello"
if ($res -ne "OK:hello") {
    Write-Host "FAIL: ValidateScript truthy return failed"
    exit 1
}
Write-Host "PASS"
exit 0
