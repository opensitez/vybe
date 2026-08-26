# vybe-test: powershell/comment_syntax_suite/block_comments
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
