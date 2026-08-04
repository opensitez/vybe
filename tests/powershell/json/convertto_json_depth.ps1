# vybe-test: powershell/json/convertto_json_depth
$deep = @{ level1 = @{ level2 = @{ level3 = "value" } } }
$json = $deep | ConvertTo-Json -Depth 5
$back = $json | ConvertFrom-Json
if ($back.level1.level2.level3 -ne "value") {
    Write-Host "FAIL: deep value not preserved"
    exit 1
}
Write-Host "PASS"
exit 0
