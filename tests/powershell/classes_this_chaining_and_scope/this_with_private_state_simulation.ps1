# vybe-test: powershell/classes_this_chaining_and_scope/this_with_private_state_simulation
class Encapsulated {
    hidden [int]$InternalState = 0
    [void]Step() { $this.InternalState += 5 }
    [int]GetState() { return $this.InternalState }
}
$enc = [Encapsulated]::new()
$enc.Step()
$enc.Step()
if ($enc.GetState() -ne 10) {
    Write-Host "FAIL: Encapsulated hidden field access via `$this failed"
    exit 1
}
Write-Host "PASS"
exit 0
