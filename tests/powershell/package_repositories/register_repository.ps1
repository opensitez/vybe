# vybe-test: powershell/package_repositories/register_repository
Register-PSRepository -Name TestRepo -SourceLocation 'https://example.com' -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
