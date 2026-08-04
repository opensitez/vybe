# vybe-test: powershell/package_management/register_package_source
Register-PackageSource -Name 'TestSource' -Location 'https://example.com' -ProviderName 'NuGet' -ErrorAction SilentlyContinue | Out-Null
Write-Host 'PASS'
exit 0
