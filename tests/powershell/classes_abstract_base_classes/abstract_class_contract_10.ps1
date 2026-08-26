# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_10
class BaseContract_10 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_10 : BaseContract_10 {
    [string]GetKind() { return "Derived_10" }
}
[BaseContract_10]$inst = [DerivedContract_10]::new()
if ($inst.GetKind() -ne "Derived_10") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
