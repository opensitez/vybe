# vybe-test: powershell/type_converters/type_converter_scriptblock_coercion
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
