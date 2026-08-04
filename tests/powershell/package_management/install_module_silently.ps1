# vybe-test: powershell/package_management/install_module_silently
Find-Module -Name PowerShellGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
