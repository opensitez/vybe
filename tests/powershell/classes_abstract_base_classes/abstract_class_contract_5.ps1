# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_5
class BaseContract_5 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_5 : BaseContract_5 {
    [string]GetKind() { return "Derived_5" }
}
[BaseContract_5]$inst = [DerivedContract_5]::new()
if ($inst.GetKind() -ne "Derived_5") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
