# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_11
class BaseContract_11 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_11 : BaseContract_11 {
    [string]GetKind() { return "Derived_11" }
}
[BaseContract_11]$inst = [DerivedContract_11]::new()
if ($inst.GetKind() -ne "Derived_11") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
