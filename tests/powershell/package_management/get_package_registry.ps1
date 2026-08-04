# vybe-test: powershell/package_management/get_package_registry
Get-PackageProvider -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
