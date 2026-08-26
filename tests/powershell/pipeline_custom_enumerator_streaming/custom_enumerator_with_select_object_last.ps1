# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_select_object_last
class LastGen : System.Collections.IEnumerable {
    [int[]]$Nums = @(1, 2, 3, 4, 5)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Nums.GetEnumerator() }
}
$lg = [LastGen]::new()
$lastTwo = @($lg | Select-Object -Last 2)
if ($lastTwo.Length -ne 2 -or $lastTwo[0] -ne 4 -or $lastTwo[1] -ne 5) {
    Write-Host "FAIL: Custom enumerator Select-Object -Last failed"
    exit 1
}
Write-Host "PASS"
exit 0
