# vybe-test: powershell/join_path_cmdlet/join_path_multi_segment_three_items
# Joining three positional segments: base, child, subchild
$res = Join-Path "opt" "apps" "server"

if ($res -notmatch "^opt[/|\\]apps[/|\\]server$") {
    Write-Host "FAIL: joining 3 segments failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
