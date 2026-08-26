# vybe-test: powershell/classes_interface_implementation/interface_type_array_assignment
class FirstDisp : System.IDisposable { [void]Dispose() {} }
class SecondDisp : System.IDisposable { [void]Dispose() {} }
[System.IDisposable[]]$arr = @([FirstDisp]::new(), [SecondDisp]::new())
if ($arr.Length -ne 2) {
    Write-Host "FAIL: Interface array assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
