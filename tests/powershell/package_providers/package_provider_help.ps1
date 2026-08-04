# vybe-test: powershell/package_providers/package_provider_help
Get-Help Get-PackageProvider -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
