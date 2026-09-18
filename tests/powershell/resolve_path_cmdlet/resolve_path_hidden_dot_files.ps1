# vybe-test: powershell/resolve_path_cmdlet/resolve_path_hidden_dot_files
# Hidden directories like .git are properly resolved to PathInfo objects
$info = Resolve-Path ".git"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path returned `$null for .git"
    exit 1
}

if ($info.Path -notmatch "\.git$") {
    Write-Host "FAIL: resolved path does not end with '.git': '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
