# vybe-test: powershell/try_finally/nested_try_finally
$ran = 0
try { try { $ran += 1 } finally { $ran += 10 } } finally { $ran += 100 }
if ($ran -eq 111) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
