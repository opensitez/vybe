# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_4
class BaseContract_4 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_4 : BaseContract_4 {
    [string]GetKind() { return "Derived_4" }
}
[BaseContract_4]$inst = [DerivedContract_4]::new()
if ($inst.GetKind() -ne "Derived_4") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
