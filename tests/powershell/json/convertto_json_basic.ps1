# vybe-test: powershell/json/convertto_json_basic
$obj = [PSCustomObject]@{ Name = "Bob"; Age = 25 }
$json = $obj | ConvertTo-Json -Compress
if ($json -notmatch '"Name"') { Write-Host "FAIL: missing Name key"; exit 1 }
if ($json -notmatch '"Bob"')  { Write-Host "FAIL: missing Bob value"; exit 1 }
if ($json -notmatch '25')     { Write-Host "FAIL: missing age 25"; exit 1 }
Write-Host "PASS"
exit 0
