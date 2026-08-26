# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_15
class BaseContract_15 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_15 : BaseContract_15 {
    [string]GetKind() { return "Derived_15" }
}
[BaseContract_15]$inst = [DerivedContract_15]::new()
if ($inst.GetKind() -ne "Derived_15") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
