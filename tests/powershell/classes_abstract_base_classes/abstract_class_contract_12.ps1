# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_12
class BaseContract_12 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_12 : BaseContract_12 {
    [string]GetKind() { return "Derived_12" }
}
[BaseContract_12]$inst = [DerivedContract_12]::new()
if ($inst.GetKind() -ne "Derived_12") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
