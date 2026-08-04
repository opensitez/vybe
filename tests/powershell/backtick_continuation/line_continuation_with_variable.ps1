# vybe-test: powershell/backtick_continuation/line_continuation_with_variable
$x = 1 + `
2
if ($x -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
