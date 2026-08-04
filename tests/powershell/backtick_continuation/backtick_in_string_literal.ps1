# vybe-test: powershell/backtick_continuation/backtick_in_string_literal
$text = "line1`nline2"
if ($text -notlike '*line1*') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
