# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_13
class BaseContract_13 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_13 : BaseContract_13 {
    [string]GetKind() { return "Derived_13" }
}
[BaseContract_13]$inst = [DerivedContract_13]::new()
if ($inst.GetKind() -ne "Derived_13") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
