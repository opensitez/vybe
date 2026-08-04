# vybe-test: powershell/json/convertfrom_json
$json = '{"Name":"Bob","Age":25}'
$obj = $json | ConvertFrom-Json
if ($obj.Name -ne "Bob") {
    Write-Host "FAIL: expected Name 'Bob', got '$($obj.Name)'"
    exit 1
}
if ($obj.Age -ne 25) {
    Write-Host "FAIL: expected Age 25, got $($obj.Age)"
    exit 1
}
Write-Host "PASS"
exit 0
