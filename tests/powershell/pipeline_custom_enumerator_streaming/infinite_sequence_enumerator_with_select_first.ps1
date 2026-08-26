# vybe-test: powershell/pipeline_custom_enumerator_streaming/infinite_sequence_enumerator_with_select_first
class NaturalNumbers : System.Collections.IEnumerable {
    [System.Collections.IEnumerator]GetEnumerator() {
        return [NatEnum]::new()
    }
}
class NatEnum : System.Collections.IEnumerator {
    [int]$Val = 0
    [object] get_Current() { return $this.Val }
    [bool] MoveNext() { $this.Val++; return $true }
    [void] Reset() { $this.Val = 0 }
}
$nats = [NaturalNumbers]::new()
$res = @($nats | Select-Object -First 5)
if ($res.Length -ne 5 -or $res[0] -ne 1 -or $res[4] -ne 5) {
    Write-Host "FAIL: Infinite sequence streaming with Select-Object -First failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
