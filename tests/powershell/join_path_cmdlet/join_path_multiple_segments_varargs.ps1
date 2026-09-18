# vybe-test: powershell/join_path_cmdlet/join_path_multiple_segments_varargs
# In PowerShell v6+, Join-Path accepts arbitrary additional positional segments
$res = Join-Path "a" "b" "c" "d"

if ($res -notmatch "^a[/|\\]b[/|\\]c[/|\\]d$") {
    Write-Host "FAIL: unexpected multi-segment path: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
