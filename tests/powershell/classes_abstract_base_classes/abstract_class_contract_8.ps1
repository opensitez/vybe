# vybe-test: powershell/classes_abstract_base_classes/abstract_class_contract_8
class BaseContract_8 {
    [string]GetKind() { return "Base" }
}
class DerivedContract_8 : BaseContract_8 {
    [string]GetKind() { return "Derived_8" }
}
[BaseContract_8]$inst = [DerivedContract_8]::new()
if ($inst.GetKind() -ne "Derived_8") { Write-Host "FAIL: Abstract base contract failed"; exit 1 }
Write-Host "PASS"; exit 0
