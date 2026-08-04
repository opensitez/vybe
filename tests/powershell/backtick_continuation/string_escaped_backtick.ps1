# vybe-test: powershell/backtick_continuation/string_escaped_backtick
$text = "Hello ``World"
if ($text -ne 'Hello `World') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
