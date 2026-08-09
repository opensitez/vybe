# vybe-test: powershell/should_process/should_process_class_method
class Processor {
    [bool] CanProcess([string]$target) {
        return $true
    }
}
$p = [Processor]::new()
if (-not $p.CanProcess("Target")) {
    Write-Host "FAIL: class method boolean evaluation failed"
    exit 1
}
Write-Host "PASS"
exit 0
