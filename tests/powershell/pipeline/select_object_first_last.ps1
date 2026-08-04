# vybe-test: powershell/pipeline/select_object_first_last
$items = 1..20
$first3 = $items | Select-Object -First 3
$last3  = $items | Select-Object -Last 3
if ($first3.Count -ne 3)  { Write-Host "FAIL: first count"; exit 1 }
if ($first3[0] -ne 1)     { Write-Host "FAIL: first[0]";    exit 1 }
if ($last3[-1] -ne 20)    { Write-Host "FAIL: last[-1]";    exit 1 }
Write-Host "PASS"
exit 0
