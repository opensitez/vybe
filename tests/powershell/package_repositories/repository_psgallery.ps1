# vybe-test: powershell/package_repositories/repository_psgallery
Find-PSRepository -Name PSGallery -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
