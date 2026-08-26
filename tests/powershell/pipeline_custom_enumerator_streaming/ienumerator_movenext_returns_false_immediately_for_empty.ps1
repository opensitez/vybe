# vybe-test: powershell/pipeline_custom_enumerator_streaming/ienumerator_movenext_returns_false_immediately_for_empty
class EmptyGen : System.Collections.IEnumerable {
    [System.Collections.IEnumerator]GetEnumerator() {
        return [EmptyEnum]::new()
    }
}
class EmptyEnum : System.Collections.IEnumerator {
    [object] get_Current() { return $null }
    [bool] MoveNext() { return $false }
    [void] Reset() {}
}
$eg = [EmptyGen]::new()
$res = @($eg | ForEach-Object { $_ })
if ($res.Length -ne 0) {
    Write-Host "FAIL: Empty custom enumerator failed"
    exit 1
}
Write-Host "PASS"
exit 0
