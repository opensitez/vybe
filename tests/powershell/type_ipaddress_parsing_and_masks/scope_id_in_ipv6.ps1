# vybe-test: powershell/type_ipaddress_parsing_and_masks/scope_id_in_ipv6
$ip = [System.Net.IPAddress]::Parse("fe80::1%4")
if ($ip.ScopeId -ne 4) {
    Write-Host "FAIL: IPv6 ScopeId extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
