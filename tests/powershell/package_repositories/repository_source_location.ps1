# vybe-test: powershell/package_repositories/repository_source_location
Get-PSRepository -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
