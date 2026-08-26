# vybe-test: powershell/classes_polymorphism_and_casting/type_test_as_operator_safe_cast
class BaseTypeItem {}
class DerivedTypeItem : BaseTypeItem {}
class UnrelatedTypeItem {}

$d = [DerivedTypeItem]::new()
$b = $d -as [BaseTypeItem]
$u = $d -as [UnrelatedTypeItem]
if ($b -eq $null -or $u -ne $null) {
    Write-Host "FAIL: -as safe casting operator failed"
    exit 1
}
Write-Host "PASS"
exit 0
