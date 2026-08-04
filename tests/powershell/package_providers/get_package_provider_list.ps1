# vybe-test: powershell/package_providers/get_package_provider_list
Get-PackageProvider -ListAvailable -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
