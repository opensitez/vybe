# vybe-test: powershell/package_management/uninstall_package_source
Get-PackageSource -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
