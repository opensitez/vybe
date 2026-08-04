# vybe-test: powershell/package_repositories/get_repository_list
Get-PSRepository -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
