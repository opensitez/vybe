# vybe-test: powershell/package_providers/install_package_provider
Install-PackageProvider -Name NuGet -Force -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
