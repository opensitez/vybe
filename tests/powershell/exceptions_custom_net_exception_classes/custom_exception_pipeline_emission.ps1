# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_pipeline_emission
class PipeException : System.Exception {
    [int]$Step
    PipeException([int]$s) : base("Failed at step $s") { $this.Step = $s }
}
$caughtStep = 0
try {
    1..5 | ForEach-Object { if ($_ -eq 3) { throw [PipeException]::new($_) } }
} catch [PipeException] {
    $caughtStep = $_.Exception.Step
}
if ($caughtStep -ne 3) {
    Write-Host "FAIL: Custom exception in pipeline failed, got $caughtStep"
    exit 1
}
Write-Host "PASS"
exit 0
