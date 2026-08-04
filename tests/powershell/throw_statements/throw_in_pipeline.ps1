# vybe-test: powershell/throw_statements/throw_in_pipeline
$thrown = $false
try { 1..3 | ForEach-Object { if ($_ -eq 2) { throw 'ERROR' } } } catch { $thrown = $true }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
