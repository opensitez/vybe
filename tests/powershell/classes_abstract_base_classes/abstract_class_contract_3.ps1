# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_3
class BaseContract_3 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_3 : BaseContract_3 {
    [string]GetKind() { return "Derived_3" }
}
[BaseContract_3]$inst = [DerivedContract_3]::new()
if ($inst.GetKind() -ne "Derived_3") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
