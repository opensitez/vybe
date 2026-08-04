# vybe-test: powershell/try_catch/catch_specific_nonmatching
try {
    throw [System.ArgumentException]::new('x')
} catch [System.InvalidOperationException] {
    Write-Host 'FAIL'
    exit 1
} catch {
    Write-Output 'CAUGHT'
}
Write-Host 'PASS'
exit 0
