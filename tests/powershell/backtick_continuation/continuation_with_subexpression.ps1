# vybe-test: powershell/backtick_continuation/continuation_with_subexpression
$value = $(1 + `
2)
if ($value -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
