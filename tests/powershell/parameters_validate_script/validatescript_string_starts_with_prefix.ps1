# vybe-test: powershell/parameters_validate_script/validatescript_string_starts_with_prefix
function Set-PrefixedCode {
    param([ValidateScript({ $_.StartsWith("APP_") })][string]$Code)
    return $Code
}
$res = Set-PrefixedCode -Code "APP_1234"
if ($res -ne "APP_1234") {
    Write-Host "FAIL: ValidateScript string prefix check failed"
    exit 1
}
Write-Host "PASS"
exit 0
