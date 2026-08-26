# vybe-test: powershell/type_ipaddress_parsing_and_masks/none_address_constant
$none = [System.Net.IPAddress]::None
if ($none.ToString() -ne "255.255.255.255") {
    Write-Host "FAIL: None address constant mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
