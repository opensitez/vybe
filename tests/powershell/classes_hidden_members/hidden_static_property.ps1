# vybe-test: powershell/classes_hidden_members/hidden_static_property
class GlobalConfig {
    hidden static [string]$ApiKey = "secret_key_abc"
    static [string]GetApiKey() {
        return [GlobalConfig]::ApiKey
    }
}
$key = [GlobalConfig]::GetApiKey()
if ($key -ne "secret_key_abc" -or [GlobalConfig]::ApiKey -ne "secret_key_abc") {
    Write-Host "FAIL: Hidden static property access failed"
    exit 1
}
Write-Host "PASS"
exit 0
