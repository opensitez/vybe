# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_select_object_skip
class SkipGen : System.Collections.IEnumerable {
    [int[]]$Items = @(10, 20, 30, 40, 50)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Items.GetEnumerator() }
}
$sg = [SkipGen]::new()
$res = @($sg | Select-Object -Skip 2 -First 2)
if ($res.Length -ne 2 -or $res[0] -ne 30 -or $res[1] -ne 40) {
    Write-Host "FAIL: Custom enumerator with Skip and First failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
