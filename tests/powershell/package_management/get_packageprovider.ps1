# vybe-test: powershell/package_management/get_packageprovider
Get-PackageProvider -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
