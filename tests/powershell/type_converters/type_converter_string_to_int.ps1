# vybe-test: powershell/type_converters/type_converter_string_to_int
$i = [System.Convert]::ChangeType("123", [int])
if ($i -ne 123 -or -not ($i -is [int])) {
    Write-Host "FAIL: ChangeType string to int expected 123"
    exit 1
}
Write-Host "PASS"
exit 0
