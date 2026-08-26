# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_nested_pipeline_consumption
class InnerStream : System.Collections.IEnumerable {
    [int[]]$V = @(1, 2)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.V.GetEnumerator() }
}
function Consume-Inner {
    process { [InnerStream]::new() }
}
$res = @(Consume-Inner | ForEach-Object { $_ })
if ($res.Length -ne 2 -or $res[0] -ne 1 -or $res[1] -ne 2) {
    Write-Host "FAIL: Nested pipeline consumption of custom enumerator failed"
    exit 1
}
Write-Host "PASS"
exit 0
