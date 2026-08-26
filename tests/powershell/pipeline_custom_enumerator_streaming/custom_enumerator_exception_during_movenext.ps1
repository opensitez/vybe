# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_exception_during_movenext
class BrokenGen : System.Collections.IEnumerable {
    [System.Collections.IEnumerator]GetEnumerator() { return [BrokenEnum]::new() }
}
class BrokenEnum : System.Collections.IEnumerator {
    [object] get_Current() { return 1 }
    [bool] MoveNext() { throw "EnumeratorBroken" }
    [void] Reset() {}
}
$bg = [BrokenGen]::new()
$caught = $false
try {
    $bg | ForEach-Object { $_ }
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when custom MoveNext throws"
    exit 1
}
Write-Host "PASS"
exit 0
