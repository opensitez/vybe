# vybe-test: powershell/parameters_validate_script/validatescript_ip_address_not_loopback
function Set-HostIP {
    param([ValidateScript({ -not [System.Net.IPAddress]::IsLoopback([System.Net.IPAddress]::Parse($_)) })][string]$IP)
    return $IP
}
$res = Set-HostIP -IP "8.8.8.8"
if ($res -ne "8.8.8.8") {
    Write-Host "FAIL: ValidateScript loopback filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
