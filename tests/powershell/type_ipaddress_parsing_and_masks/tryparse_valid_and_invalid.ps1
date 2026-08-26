# vybe-test: powershell/type_ipaddress_parsing_and_masks/tryparse_valid_and_invalid
$ip = $null
$ok = [System.Net.IPAddress]::TryParse("10.0.0.1", [ref]$ip)
$bad = [System.Net.IPAddress]::TryParse("999.999.999.999", [ref]$ip)
if (-not $ok -or $bad) {
    Write-Host "FAIL: IPAddress TryParse check failed"
    exit 1
}
Write-Host "PASS"
exit 0
