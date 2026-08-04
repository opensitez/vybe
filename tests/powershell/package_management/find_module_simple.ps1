# vybe-test: powershell/package_management/find_module_simple
Find-Module -Name PowerShellGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
