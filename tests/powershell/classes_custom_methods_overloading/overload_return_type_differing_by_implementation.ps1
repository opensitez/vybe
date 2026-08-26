# vybe-test: powershell/classes_custom_methods_overloading/overload_return_type_differing_by_implementation
class ConverterClass {
    [int]ConvertVal([string]$s) { return [int]::Parse($s) }
    [string]ConvertVal([int]$i) { return "Number-$i" }
}
$cc = [ConverterClass]::new()
$i = $cc.ConvertVal("123")
$s = $cc.ConvertVal(456)
if ($i -ne 123 -or $s -ne "Number-456") {
    Write-Host "FAIL: Overload differing return types failed"
    exit 1
}
Write-Host "PASS"
exit 0
