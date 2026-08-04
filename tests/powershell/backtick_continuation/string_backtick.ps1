# vybe-test: powershell/backtick_continuation/string_backtick
$text = "Hello `
World"
if ($text -notlike '*Hello*') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
