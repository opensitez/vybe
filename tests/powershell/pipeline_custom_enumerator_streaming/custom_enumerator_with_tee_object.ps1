# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_tee_object
class DataStream : System.Collections.IEnumerable {
    [int[]]$Vals = @(10, 20)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Vals.GetEnumerator() }
}
$ds = [DataStream]::new()
$sideBuffer = [System.Collections.Generic.List[int]]::new()
$res = @($ds | Tee-Object -Variable sideBuffer)
if ($res.Length -ne 2 -or $sideBuffer.Count -ne 2) {
    Write-Host "FAIL: Custom enumerator with Tee-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
