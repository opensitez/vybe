# vybe-test: powershell/enums/enum_basic
enum Direction { North; South; East; West }
$d = [Direction]::North
if ($d -ne [Direction]::North) { Write-Host "FAIL: identity"; exit 1 }
if ($d -eq [Direction]::South) { Write-Host "FAIL: inequality"; exit 1 }
$name = $d.ToString()
if ($name -ne "North") { Write-Host "FAIL: ToString '$name'"; exit 1 }
Write-Host "PASS"
exit 0
