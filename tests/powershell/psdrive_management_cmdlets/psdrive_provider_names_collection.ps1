# vybe-test: powershell/psdrive_management_cmdlets/psdrive_provider_names_collection
# The provider names of mounted PSDrives include standard PowerShell providers
$providers = @((Get-PSDrive).Provider.Name | Select-Object -Unique)

if ($providers -notcontains "Environment") {
    Write-Host "FAIL: 'Environment' provider missing from mounted drives"
    exit 1
}

if ($providers -notcontains "Variable") {
    Write-Host "FAIL: 'Variable' provider missing from mounted drives"
    exit 1
}

if ($providers -notcontains "FileSystem") {
    Write-Host "FAIL: 'FileSystem' provider missing from mounted drives"
    exit 1
}

Write-Host "PASS"
exit 0
