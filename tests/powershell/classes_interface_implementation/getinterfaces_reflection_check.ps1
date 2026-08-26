# vybe-test: powershell/classes_interface_implementation/getinterfaces_reflection_check
class CheckInter : System.IDisposable {
    [void]Dispose() {}
}
$interfaces = @([CheckInter].GetInterfaces() | ForEach-Object { $_.Name })
if (-not ($interfaces -contains "IDisposable")) {
    Write-Host "FAIL: GetInterfaces reflection check failed"
    exit 1
}
Write-Host "PASS"
exit 0
