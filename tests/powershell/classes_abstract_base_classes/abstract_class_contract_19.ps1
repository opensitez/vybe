# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_19
class BaseContract_19 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_19 : BaseContract_19 {
    [string]GetKind() { return "Derived_19" }
}
[BaseContract_19]$inst = [DerivedContract_19]::new()
if ($inst.GetKind() -ne "Derived_19") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
