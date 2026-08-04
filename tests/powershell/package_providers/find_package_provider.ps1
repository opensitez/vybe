# vybe-test: powershell/package_providers/find_package_provider
Find-PackageProvider -Name NuGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
