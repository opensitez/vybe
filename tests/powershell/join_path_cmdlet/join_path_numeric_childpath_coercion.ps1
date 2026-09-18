# vybe-test: powershell/join_path_cmdlet/join_path_numeric_childpath_coercion
# Numeric arguments passed as child paths are automatically coerced to string path segments
$res = Join-Path "versions" 2 5

if ($res -notmatch "^versions[/|\\]2[/|\\]5$") {
    Write-Host "FAIL: numeric child path coercion failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
