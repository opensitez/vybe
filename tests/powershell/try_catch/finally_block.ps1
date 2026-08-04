# vybe-test: powershell/try_catch/finally_block
$flag = $false
try {
    throw 'err'
} catch {
    $flag = $true
} finally {
    if (-not $flag) { Write-Host 'FAIL'; exit 1 }
}
Write-Host 'PASS'
exit 0
