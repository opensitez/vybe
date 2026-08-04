# vybe-test: powershell/interpolation/here_string_interpolation
$name = 'Hi'
$text = @"
$name
"@
if ($text -notlike '*Hi*') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
