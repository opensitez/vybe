# vybe-test: powershell/join_path_cmdlet/join_path_relative_parent_dot_dot
# Joining parent directory '..' with sibling segment preserves relative parent navigation
$res = Join-Path ".." "sibling"

if ($res -notmatch "^\.\.[/|\\]sibling$") {
    Write-Host "FAIL: unexpected parent relative joined path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
