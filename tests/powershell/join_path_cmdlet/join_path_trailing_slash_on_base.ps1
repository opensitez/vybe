# vybe-test: powershell/join_path_cmdlet/join_path_trailing_slash_on_base
# Join-Path normalizes trailing slashes on base path without introducing double slashes
$res = Join-Path "dir/" "file.txt"

if ($res -notmatch "^dir[/|\\]file\.txt$") {
    Write-Host "FAIL: trailing slash on base path produced invalid joined path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
