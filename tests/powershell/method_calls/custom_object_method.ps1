# vybe-test: powershell/method_calls/custom_object_method
$obj = [pscustomobject]@{ Value = 1 }
if ($obj.ToString() -ne $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
