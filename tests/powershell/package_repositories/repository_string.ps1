# vybe-test: powershell/package_repositories/repository_string
if ('Test' -is [string]) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
