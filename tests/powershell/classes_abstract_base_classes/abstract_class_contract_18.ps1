# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_18
class BaseContract_18 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_18 : BaseContract_18 {
    [string]GetKind() { return "Derived_18" }
}
[BaseContract_18]$inst = [DerivedContract_18]::new()
if ($inst.GetKind() -ne "Derived_18") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
