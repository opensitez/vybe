# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_reset_support
class ResettableEnum : System.Collections.IEnumerator {
    [int[]]$Data = @(100, 200)
    [int]$Idx = -1
    [object] get_Current() { return $this.Data[$this.Idx] }
    [bool] MoveNext() { $this.Idx++; return ($this.Idx -lt $this.Data.Length) }
    [void] Reset() { $this.Idx = -1 }
}
$e = [ResettableEnum]::new()
$null = $e.MoveNext()
$val1 = $e.Current
$e.Reset()
$null = $e.MoveNext()
$val2 = $e.Current
if ($val1 -ne 100 -or $val2 -ne 100) {
    Write-Host "FAIL: Custom enumerator Reset method failed"
    exit 1
}
Write-Host "PASS"
exit 0
