# vybe-test: powershell/try_catch/catch_variable
try {
    throw 'err'
} catch {
    if ($_ -eq $null) { Write-Host 'FAIL'; exit 1 }
}
Write-Host 'PASS'
exit 0
