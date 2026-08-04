# vybe-test: powershell/json/convertfrom_json_nested
$json = '{"user":{"name":"Carol","roles":["admin","editor"]}}'
$obj = $json | ConvertFrom-Json
if ($obj.user.name -ne "Carol") { Write-Host "FAIL: name"; exit 1 }
if ($obj.user.roles[0] -ne "admin")  { Write-Host "FAIL: roles[0]"; exit 1 }
if ($obj.user.roles[1] -ne "editor") { Write-Host "FAIL: roles[1]"; exit 1 }
Write-Host "PASS"
exit 0
