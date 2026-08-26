# vybe-test: powershell/parameters_validate_pattern/validatepattern_with_splatted_hashtable
function Set-ApiKey {
    param([ValidatePattern('^key_[a-f0-9]{16}$')][string]$Key)
    return "KeyValid"
}
$params = @{ Key = "key_0123456789abcdef" }
$res = Set-ApiKey @params
if ($res -ne "KeyValid") {
    Write-Host "FAIL: ValidatePattern splatting failed"
    exit 1
}
Write-Host "PASS"
exit 0
