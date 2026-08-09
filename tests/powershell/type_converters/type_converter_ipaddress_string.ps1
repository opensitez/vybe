# vybe-test: powershell/type_converters/type_converter_ipaddress_string
$ip = [ipaddress]"127.0.0.1"
if ($ip.IPAddressToString -ne "127.0.0.1") {
    Write-Host "FAIL: string to [ipaddress] conversion expected 127.0.0.1"
    exit 1
}
Write-Host "PASS"
exit 0
