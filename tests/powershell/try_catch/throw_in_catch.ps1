# vybe-test: powershell/try_catch/throw_in_catch
try {
    throw 'err'
} catch {
    throw 'new'
}
catch {
    Write-Output 'CAUGHT'
}
Write-Host 'PASS'
exit 0
