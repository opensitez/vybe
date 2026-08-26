# vybe-test: powershell/type_ipaddress_parsing_and_masks/loopback_ipv4_and_ipv6
$v4Loop = [System.Net.IPAddress]::Loopback
$v6Loop = [System.Net.IPAddress]::IPv6Loopback
if ($v4Loop.ToString() -ne "127.0.0.1" -or $v6Loop.ToString() -ne "::1") {
    Write-Host "FAIL: Static loopback constants mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
