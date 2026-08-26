# vybe-test: powershell/type_ipaddress_parsing_and_masks/ip_in_hashtable_lookup
$ip = [System.Net.IPAddress]::Parse("1.1.1.1")
$ht = @{ $ip = "Cloudflare" }
if ($ht[$ip] -ne "Cloudflare") {
    Write-Host "FAIL: IPAddress hashtable key lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
