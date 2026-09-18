# vybe-test: powershell/join_path_cmdlet/join_path_leading_slash_on_child
# Join-Path normalizes leading slashes on child path without introducing duplicate slashes
$res = Join-Path "dir" "/file.txt"

if ($res -notmatch "^dir[/|\\]file\.txt$") {
    Write-Host "FAIL: leading slash on child path produced invalid joined path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
