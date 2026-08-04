# vybe-test: powershell/interpolation/basic_interpolation
$name = 'Vybe'
$text = "Hello, $name"
if ($text -ne 'Hello, Vybe') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
