# vybe-test: powershell/literals/expanded_string_literal
$name = 'Vybe'
$text = "Hello $name"
if ($text -ne 'Hello Vybe') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
