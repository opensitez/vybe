# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_2
class BaseContract_2 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_2 : BaseContract_2 {
    [string]GetKind() { return "Derived_2" }
}
[BaseContract_2]$inst = [DerivedContract_2]::new()
if ($inst.GetKind() -ne "Derived_2") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
