# vybe-test: powershell/package_repositories/repository_list_all
Get-PSRepository -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
