# vybe-test: powershell/automatic_variables/dollar_null_behavior
$x = $null
if ($null -ne $x) { Write-Host "FAIL: should be null"; exit 1 }
if ($x -ne $null) { Write-Host "FAIL: symmetric comparison"; exit 1 }
$arr = @(1, $null, 3)
$count = ($arr | Where-Object { $_ -ne $null } | Measure-Object).Count
if ($count -ne 2) { Write-Host "FAIL: filter null from array"; exit 1 }
Write-Host "PASS"
exit 0
