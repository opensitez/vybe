# vybe-test: powershell/join_path_cmdlet/join_path_relative_dot_segments
# Joining current directory '.' with child segment preserves relative dot path
$res = Join-Path "." "subdir"

if ($res -notmatch "^\.[/|\\]subdir$") {
    Write-Host "FAIL: unexpected relative joined path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
