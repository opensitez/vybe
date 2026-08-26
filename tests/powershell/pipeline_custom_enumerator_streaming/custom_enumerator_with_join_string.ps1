# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_join_string
class CharStream : System.Collections.IEnumerable {
    [char[]]$Chars = @('H', 'i')
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Chars.GetEnumerator() }
}
$cs = [CharStream]::new()
$res = -join @($cs)
if ($res -ne "Hi") {
    Write-Host "FAIL: Custom enumerator join failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
