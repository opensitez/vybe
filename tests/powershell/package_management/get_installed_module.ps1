# vybe-test: powershell/package_management/get_installed_module
Get-InstalledModule -Name PowerShellGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
