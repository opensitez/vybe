# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_1
class BaseContract_1 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_1 : BaseContract_1 {
    [string]GetKind() { return "Derived_1" }
}
[BaseContract_1]$inst = [DerivedContract_1]::new()
if ($inst.GetKind() -ne "Derived_1") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
