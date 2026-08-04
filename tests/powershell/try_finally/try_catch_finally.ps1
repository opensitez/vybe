# vybe-test: powershell/try_finally/try_catch_finally
$ran = $false
try { throw 'ERR' } catch { $ran = $true } finally { if ($ran) { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
