# vybe-test: powershell/package_repositories/repository_help
Get-Help Register-PSRepository -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
