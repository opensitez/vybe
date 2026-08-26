# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_6
class BaseContract_6 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_6 : BaseContract_6 {
    [string]GetKind() { return "Derived_6" }
}
[BaseContract_6]$inst = [DerivedContract_6]::new()
if ($inst.GetKind() -ne "Derived_6") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
