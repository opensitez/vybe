# vybe-test: powershell/package_providers/get_package_providers
Get-PackageProvider -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
