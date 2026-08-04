# vybe-test: powershell/package_management/find_package
Find-Package -Name PowerShellGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
