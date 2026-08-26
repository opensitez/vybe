# vybe-test: powershell/parameters_validate_pattern/validatepattern_ip_address_format
function Set-IP {
    param([ValidatePattern('^(\d{1,3}\.){3}\d{1,3}$')][string]$IP)
    return $IP
}
$res = Set-IP -IP "192.168.1.1"
if ($res -ne "192.168.1.1") {
    Write-Host "FAIL: ValidatePattern IP failed"
    exit 1
}
Write-Host "PASS"
exit 0
