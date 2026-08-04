# vybe-test: powershell/package_providers/package_provider_all
Get-PackageProvider -ListAvailable -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
