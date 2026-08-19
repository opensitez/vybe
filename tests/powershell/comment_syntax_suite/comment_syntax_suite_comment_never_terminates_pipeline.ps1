# vybe-test: powershell/comment_syntax_suite/comment_never_terminates_pipeline
$sum = 1,2,3 | Measure-Object -Sum # comment at end of pipeline line

if ($sum.Count -ne 1 -or $sum.Sum -ne 6) {
    Write-Host "FAIL: pipeline was terminated by comment: sum=$($sum.Sum)"
    exit 1
}

Write-Host 'PASS'
exit 0
