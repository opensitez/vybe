# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_14
class BaseContract_14 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_14 : BaseContract_14 {
    [string]GetKind() { return "Derived_14" }
}
[BaseContract_14]$inst = [DerivedContract_14]::new()
if ($inst.GetKind() -ne "Derived_14") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
