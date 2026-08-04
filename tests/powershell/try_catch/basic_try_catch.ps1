# vybe-test: powershell/try_catch/basic_try_catch
try {
    throw 'error'
} catch {
    Write-Output 'CAUGHT'
}
Write-Host 'PASS'
exit 0
