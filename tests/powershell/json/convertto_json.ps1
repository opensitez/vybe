# vybe-test: powershell/json/convertto_json
$obj = @{ Name = "Alice"; Age = 30 }
$json = $obj | ConvertTo-Json -Compress
if ($json -notmatch '"Name":"Alice"') {
    Write-Host "FAIL: expected JSON with Name:Alice, got '$json'"
    exit 1
}
Write-Host "PASS"
exit 0
