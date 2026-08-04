# vybe-test: powershell/try_catch/nested_try_catch
try {
    try { throw 'err' } catch { }
} catch {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
