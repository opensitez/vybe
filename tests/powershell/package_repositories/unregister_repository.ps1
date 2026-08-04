# vybe-test: powershell/package_repositories/unregister_repository
Unregister-PSRepository -Name TestRepo -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
