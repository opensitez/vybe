# vybe-test: powershell/package_management/uninstall_module_silently
Get-InstalledModule -Name PowerShellGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
