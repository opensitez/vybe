# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_ienumerable_class_pipeline_streaming
class RangeGenerator : System.Collections.IEnumerable {
    [int]$Start; [int]$Count
    RangeGenerator([int]$s, [int]$c) { $this.Start = $s; $this.Count = $c }
    [System.Collections.IEnumerator]GetEnumerator() {
        return [RangeEnumerator]::new($this.Start, $this.Count)
    }
}
class RangeEnumerator : System.Collections.IEnumerator {
    [int]$Start; [int]$Count; [int]$CurrentIdx = -1
    RangeEnumerator([int]$s, [int]$c) { $this.Start = $s; $this.Count = $c }
    [object] get_Current() { return $this.Start + $this.CurrentIdx }
    [bool] MoveNext() {
        $this.CurrentIdx++
        return ($this.CurrentIdx -lt $this.Count)
    }
    [void] Reset() { $this.CurrentIdx = -1 }
}
$gen = [RangeGenerator]::new(10, 3) # 10, 11, 12
$res = @($gen | ForEach-Object { $_ * 2 })
if ($res.Length -ne 3 -or $res[0] -ne 20 -or $res[1] -ne 22 -or $res[2] -ne 24) {
    Write-Host "FAIL: Custom IEnumerable streaming in pipeline failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
