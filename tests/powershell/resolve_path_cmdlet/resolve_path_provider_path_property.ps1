# vybe-test: powershell/resolve_path_cmdlet/resolve_path_provider_path_property
# The .ProviderPath property of PathInfo exposes the provider-internal canonical path
$info = Resolve-Path "tests"

if ($null -eq $info.ProviderPath) {
    Write-Host "FAIL: ProviderPath property is `$null"
    exit 1
}

if ($info.ProviderPath -ne $info.Path) {
    Write-Host "FAIL: ProviderPath ($($info.ProviderPath)) did not match Path ($($info.Path)) for FileSystem"
    exit 1
}

Write-Host "PASS"
exit 0
