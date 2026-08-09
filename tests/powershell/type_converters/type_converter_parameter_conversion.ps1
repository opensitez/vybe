# vybe-test: powershell/type_converters/type_converter_parameter_conversion
function Set-TargetIp {
    param([ipaddress]$IP)
    return $IP.IPAddressToString
}
$res = Set-TargetIp "10.0.0.1"
if ($res -ne "10.0.0.1") {
    Write-Host "FAIL: parameter ipaddress conversion expected 10.0.0.1, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
