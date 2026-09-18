# vybe-test: powershell/resolve_path_cmdlet/resolve_path_environment_psdrive
# Resolve-Path against an Environment PSDrive path resolves to the 'Environment' provider
$info = Resolve-Path "Env:PATH"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path returned `$null for Env:PATH"
    exit 1
}

if ($info.Provider.Name -ne "Environment") {
    Write-Host "FAIL: expected provider 'Environment', got: '$($info.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
