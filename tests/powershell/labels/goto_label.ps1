# vybe-test: powershell/labels/goto_label
$flag = $false
:label
$flag = $true
if ($flag) { Write-Output 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
