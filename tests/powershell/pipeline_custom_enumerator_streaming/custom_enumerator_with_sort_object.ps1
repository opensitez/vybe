# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_sort_object
class UnsortedList : System.Collections.IEnumerable {
    [int[]]$Nums = @(5, 1, 4, 2, 3)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Nums.GetEnumerator() }
}
$ul = [UnsortedList]::new()
$sorted = @($ul | Sort-Object)
if ($sorted[0] -ne 1 -or $sorted[4] -ne 5) {
    Write-Host "FAIL: Custom enumerator Sort-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
