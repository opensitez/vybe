# vybe-test: powershell/parameters_validate_script/validatescript_with_hashtable_parameter
function Validate-ConfigTable {
    param([ValidateScript({ $_.ContainsKey("host") })][hashtable]$Config)
    return $Config["host"]
}
$res = Validate-ConfigTable -Config @{ host = "127.0.0.1"; port = 8080 }
if ($res -ne "127.0.0.1") {
    Write-Host "FAIL: ValidateScript on hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
