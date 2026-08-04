# vybe-test: powershell/try_finally/try_finally_block
$ran = 0
try { $ran += 1 } finally { $ran += 2 }
if ($ran -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
