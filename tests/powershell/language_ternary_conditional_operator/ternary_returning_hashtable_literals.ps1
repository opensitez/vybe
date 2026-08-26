# vybe-test: powershell/language_ternary_conditional_operator/ternary_returning_hashtable_literals
$isProd = $false
$config = $isProd ? @{ env = "prod" } : @{ env = "dev" }
if ($config["env"] -ne "dev") {
    Write-Host "FAIL: Ternary returning hashtable literal failed"
    exit 1
}
Write-Host "PASS"
exit 0
