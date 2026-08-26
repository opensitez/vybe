# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_streaming_large_dataset_lazily
class LazyCounter : System.Collections.IEnumerable {
    [int]$Limit
    LazyCounter([int]$l) { $this.Limit = $l }
    [System.Collections.IEnumerator]GetEnumerator() { return [LazyCounterEnum]::new($this.Limit) }
}
class LazyCounterEnum : System.Collections.IEnumerator {
    [int]$Max; [int]$Cur = 0
    LazyCounterEnum([int]$m) { $this.Max = $m }
    [object] get_Current() { return $this.Cur }
    [bool] MoveNext() { $this.Cur++; return ($this.Cur -le $this.Max) }
    [void] Reset() { $this.Cur = 0 }
}
$lc = [LazyCounter]::new(100000)
# Select only first 3 items without enumerating 100,000
$res = @($lc | Select-Object -First 3)
if ($res.Length -ne 3 -or $res[0] -ne 1 -or $res[2] -ne 3) {
    Write-Host "FAIL: Lazy large dataset streaming failed"
    exit 1
}
Write-Host "PASS"
exit 0
