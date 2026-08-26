# vybe-test: powershell/classes_interface_implementation/interface_contract_return_type_validation
class StringSupplier {
    [string]GetMessage() { return "OK" }
}
$ss = [StringSupplier]::new()
if ($ss.GetMessage() -ne "OK") {
    Write-Host "FAIL: Method contract return failed"
    exit 1
}
Write-Host "PASS"
exit 0
