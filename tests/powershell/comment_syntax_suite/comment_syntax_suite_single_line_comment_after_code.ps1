# vybe-test: powershell/comment_syntax_suite/comment_syntax_suite_single_line_comment_after_code
$values = @(
    2,
    4,
    6
)

$sum = ($values | Measure-Object -Sum).Sum  # base sum before mutation
$sum += 3

if ($sum -ne 15) {
    Write-Host "FAIL: expected 15, got $sum"
    exit 1
}

Write-Host 'PASS'
exit 0
