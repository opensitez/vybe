# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_16
class BaseContract_16 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_16 : BaseContract_16 {
    [string]GetKind() { return "Derived_16" }
}
[BaseContract_16]$inst = [DerivedContract_16]::new()
if ($inst.GetKind() -ne "Derived_16") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
