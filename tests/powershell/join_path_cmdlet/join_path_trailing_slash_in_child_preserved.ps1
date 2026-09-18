# vybe-test: powershell/join_path_cmdlet/join_path_trailing_slash_in_child_preserved
# When the final child path component has a trailing slash, Join-Path preserves it
$res = Join-Path "parent" "child/"

if ($res -notmatch "^parent[/|\\]child[/|\\]$") {
    Write-Host "FAIL: trailing slash in child component was not preserved: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
