# vybe-test: powershell/try_finally/finally_after_catch
$ran = $false
try { throw 'ERR' } catch { $ran = $true } finally { if ($ran) { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
