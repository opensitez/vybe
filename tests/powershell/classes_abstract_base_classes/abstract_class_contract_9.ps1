# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_9
class BaseContract_9 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_9 : BaseContract_9 {
    [string]GetKind() { return "Derived_9" }
}
[BaseContract_9]$inst = [DerivedContract_9]::new()
if ($inst.GetKind() -ne "Derived_9") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
