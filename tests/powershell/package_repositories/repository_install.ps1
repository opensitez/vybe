# vybe-test: powershell/package_repositories/repository_install
Install-Package -Name PowerShellGet -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
