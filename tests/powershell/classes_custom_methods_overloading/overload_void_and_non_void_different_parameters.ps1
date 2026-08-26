# vybe-test: powershell/classes_custom_methods_overloading/overload_void_and_non_void_different_parameters
class StateChanger {
    [int]$Val = 0
    [void]Update() { $this.Val = 10 }
    [int]Update([int]$newVal) { $this.Val = $newVal; return $this.Val }
}
$sc = [StateChanger]::new()
$sc.Update()
$res = $sc.Update(25)
if ($sc.Val -ne 25 -or $res -ne 25) {
    Write-Host "FAIL: Void vs non-void overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
