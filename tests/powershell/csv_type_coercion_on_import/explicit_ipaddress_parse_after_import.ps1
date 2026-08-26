# vybe-test: powershell/csv_type_coercion_on_import/explicit_ipaddress_parse_after_import
$csv = @"
Host,IP
Gateway,192.168.1.1
"@
$row = $csv | ConvertFrom-Csv
$ip = [System.Net.IPAddress]::Parse($row.IP)
if ($ip.ToString() -ne "192.168.1.1") {
    Write-Host "FAIL: Explicit IPAddress parse after import failed"
    exit 1
}
Write-Host "PASS"
exit 0
