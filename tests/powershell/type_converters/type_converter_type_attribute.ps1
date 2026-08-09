# vybe-test: powershell/type_converters/type_converter_type_attribute
function Set-Age {
    param([int]$Age)
    return $Age
}
$res = Set-Age "35"
if ($res -ne 35 -or -not ($res -is [int])) {
    Write-Host "FAIL: parameter type converter string to int expected 35"
    exit 1
}
Write-Host "PASS"
exit 0
