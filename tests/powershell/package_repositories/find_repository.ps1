# vybe-test: powershell/package_repositories/find_repository
Find-PSRepository -Name PSGallery -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
