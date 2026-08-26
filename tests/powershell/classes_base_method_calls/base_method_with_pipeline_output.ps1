# vybe-test: powershell/classes_base_method_calls/base_method_with_pipeline_output
class BasePipe {
    [int[]]GetNums() { return @(1, 2, 3) }
}
class SubPipe : BasePipe {
    [int]GetSum() {
        $nums = ([BasePipe]$this).GetNums()
        $sum = 0
        foreach ($n in $nums) { $sum += $n }
        return $sum
    }
}
$sp = [SubPipe]::new()
if ($sp.GetSum() -ne 6) {
    Write-Host "FAIL: Base method with pipeline output sum failed"
    exit 1
}
Write-Host "PASS"
exit 0
