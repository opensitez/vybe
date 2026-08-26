# vybe-test: powershell/classes_base_method_calls/base_method_invoked_inside_loop
class BaseStepper {
    [int]$Steps = 0
    [void]Step() { $this.Steps++ }
}
class MultiStepper : BaseStepper {
    [void]StepMany([int]$n) {
        for ($i = 0; $i -lt $n; $i++) {
            ([BaseStepper]$this).Step()
        }
    }
}
$ms = [MultiStepper]::new()
$ms.StepMany(5)
if ($ms.Steps -ne 5) {
    Write-Host "FAIL: Base method inside loop failed, got $($ms.Steps)"
    exit 1
}
Write-Host "PASS"
exit 0
