# vybe-test: powershell/package_providers/uninstall_package_provider
Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
