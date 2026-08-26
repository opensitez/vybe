# vybe-test: powershell/parameters_validate_script/validatescript_even_number_check_success
function Set-EvenNumber {
    param([ValidateScript({ $_ % 2 -eq 0 })][int]$Number)
    return "Even:$Number"
}
$res = Set-EvenNumber -Number 42
if ($res -ne "Even:42") {
    Write-Host "FAIL: ValidateScript even number check failed"
    exit 1
}
Write-Host "PASS"
exit 0
