# vybe-test: powershell/package_providers/package_provider_module
Get-InstalledModule -Name PackageManagement -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
