# vybe-test: powershell/throw_statements/throw_in_function
$thrown = $false
function Test-Func { throw 'ERROR' }
try { Test-Func } catch { $thrown = $true }
if ($thrown) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
