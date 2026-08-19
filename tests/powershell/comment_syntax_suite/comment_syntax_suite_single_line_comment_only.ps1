# vybe-test: powershell/comment_syntax_suite/comment_syntax_suite_single_line_comment_only
$values = @()

$values += 2
# This trailing text must be ignored by the parser.
$values += 3

if ($values.Count -ne 2) {
    Write-Host "FAIL: expected 2 values, got $($values.Count)"
    exit 1
}

if (($values | Measure-Object -Sum).Sum -ne 5) {
    Write-Host "FAIL: expected sum 5, got $(($values | Measure-Object -Sum).Sum)"
    exit 1
}

Write-Host 'PASS'
exit 0
