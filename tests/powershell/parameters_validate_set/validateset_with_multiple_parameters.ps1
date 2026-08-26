# vybe-test: powershell/parameters_validate_set/validateset_with_multiple_parameters
function Config-Service {
    param(
        [ValidateSet("HTTP", "HTTPS")][string]$Protocol,
        [ValidateSet("JSON", "XML")][string]$Format
    )
    return "$Protocol-$Format"
}
$res = Config-Service -Protocol "HTTPS" -Format "JSON"
if ($res -ne "HTTPS-JSON") {
    Write-Host "FAIL: Multiple ValidateSet parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
