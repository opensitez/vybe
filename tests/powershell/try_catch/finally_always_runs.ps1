# vybe-test: powershell/try_catch/finally_always_runs
$ran = $false
try {
    throw 'err'
} catch {
} finally {
    $ran = $true
}
if (-not $ran) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
