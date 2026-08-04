# vybe-test: powershell/try_catch/catch_exception_type
try {
    throw [System.InvalidOperationException]::new('msg')
} catch [System.InvalidOperationException] {
    Write-Output 'CAUGHT'
}
Write-Host 'PASS'
exit 0
