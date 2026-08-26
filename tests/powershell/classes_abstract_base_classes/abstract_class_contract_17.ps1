# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_17
class BaseContract_17 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_17 : BaseContract_17 {
    [string]GetKind() { return "Derived_17" }
}
[BaseContract_17]$inst = [DerivedContract_17]::new()
if ($inst.GetKind() -ne "Derived_17") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
