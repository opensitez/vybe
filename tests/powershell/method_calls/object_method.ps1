# vybe-test: powershell/method_calls/object_method
$obj = New-Object System.Text.StringBuilder
$obj.Append('x') | Out-Null
if ($obj.ToString() -eq 'x') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
