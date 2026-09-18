# vybe-test: powershell/resolve_path_cmdlet/resolve_path_returns_pathinfo_type
# Resolve-Path returns an instance of System.Management.Automation.PathInfo
$info = Resolve-Path "."

if (-not ($info -is [System.Management.Automation.PathInfo])) {
    Write-Host "FAIL: expected System.Management.Automation.PathInfo, got: $($info.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
