# vybe-test: powershell/type_ipaddress_parsing_and_masks/any_and_broadcast_addresses
$any = [System.Net.IPAddress]::Any
$bcast = [System.Net.IPAddress]::Broadcast
if ($any.ToString() -ne "0.0.0.0" -or $bcast.ToString() -ne "255.255.255.255") {
    Write-Host "FAIL: Any/Broadcast static constants mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
