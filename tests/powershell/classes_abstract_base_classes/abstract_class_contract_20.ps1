# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_20
class BaseContract_20 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_20 : BaseContract_20 {
    [string]GetKind() { return "Derived_20" }
}
[BaseContract_20]$inst = [DerivedContract_20]::new()
if ($inst.GetKind() -ne "Derived_20") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
