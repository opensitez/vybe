# vybe-test: powershell/try_catch/catch_all
try {
    throw 'err'
} catch {
    Write-Output 'CAUGHT'
}
Write-Host 'PASS'
exit 0
