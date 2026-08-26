# vybe-test: powershell/parameters_validate_script/validatescript_with_splatted_hashtable
function Set-SplattedVal {
    param([ValidateScript({ $_.Length -eq 5 })][string]$Code)
    return "Splat:$Code"
}
$params = @{ Code = "ABCDE" }
$res = Set-SplattedVal @params
if ($res -ne "Splat:ABCDE") {
    Write-Host "FAIL: ValidateScript splatting failed"
    exit 1
}
Write-Host "PASS"
exit 0
