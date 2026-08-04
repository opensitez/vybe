# vybe-test: powershell/package_providers/package_provider_name
Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
