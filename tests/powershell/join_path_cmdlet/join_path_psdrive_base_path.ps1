# vybe-test: powershell/join_path_cmdlet/join_path_psdrive_base_path
# Joining onto a PSDrive base path (e.g. Env:\) produces a valid provider path
$res = Join-Path "Env:\" "PATH"

if ($res -notmatch "^Env:[/|\\]PATH$") {
    Write-Host "FAIL: unexpected PSDrive joined path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
