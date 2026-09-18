# vybe-test: powershell/join_path_cmdlet/join_path_root_path_as_base
# Joining a root path '/' with a child path segment produces an absolute root path
$res = Join-Path "/" "usr"

if ($res -notmatch "^[/|\\]usr$") {
    Write-Host "FAIL: joining on root '/' produced invalid path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
