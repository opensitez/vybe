# vybe-test: powershell/interpolation/variable_expansion_in_string
$name = 'PS'
$text = "Hello $name"
if ($text -ne 'Hello PS') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
