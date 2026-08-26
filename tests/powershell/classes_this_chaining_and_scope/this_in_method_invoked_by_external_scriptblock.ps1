# vybe-test: powershell/classes_this_chaining_and_scope/this_in_method_invoked_by_external_scriptblock
class Invokable {
    [int]$Val = 42
    [int]GetVal() { return $this.Val }
}
$inv = [Invokable]::new()
$sb = { param($target) $target.GetVal() }
$res = & $sb $inv
if ($res -ne 42) {
    Write-Host "FAIL: External scriptblock calling `$this method failed"
    exit 1
}
Write-Host "PASS"
exit 0
