# vybe-test: powershell/comment_syntax_suite/comment_syntax_suite_comment_never_terminates_pipeline
$sum = 1 + <# comment #> 2 + <# comment #> 3
if ($sum -ne 6) {
    Write-Host "FAIL: Pipeline inline comment failed"
    exit 1
}
Write-Host "PASS"
exit 0
