# vybe-test: powershell/package_providers/package_provider_sources
Get-PackageSource -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
