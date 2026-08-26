# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_7
class BaseContract_7 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_7 : BaseContract_7 {
    [string]GetKind() { return "Derived_7" }
}
[BaseContract_7]$inst = [DerivedContract_7]::new()
if ($inst.GetKind() -ne "Derived_7") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
