# vybe-test: powershell/join_path_cmdlet/join_path_two_segments
# Join-Path combines parent directory and child filename with system path separator
$res = Join-Path "dir" "file.txt"

if ($res -notmatch "^dir[/|\\]file\.txt$") {
    Write-Host "FAIL: unexpected joined path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
