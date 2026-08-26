# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_generic_list_of_base_types
class BaseEntry { [string]$Key; BaseEntry([string]$k) { $this.Key = $k } }
class StringEntry : BaseEntry { StringEntry([string]$k) : base($k) {} }
class NumberEntry : BaseEntry { NumberEntry([string]$k) : base($k) {} }
$list = [System.Collections.Generic.List[BaseEntry]]::new()
$list.Add([StringEntry]::new("str1"))
$list.Add([NumberEntry]::new("num1"))
if ($list.Count -ne 2 -or $list[0].Key -ne "str1" -or $list[1].Key -ne "num1") {
    Write-Host "FAIL: List of base types failed"
    exit 1
}
Write-Host "PASS"
exit 0
