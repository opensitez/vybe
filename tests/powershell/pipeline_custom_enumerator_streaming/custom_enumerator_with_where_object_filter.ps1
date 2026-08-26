# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_where_object_filter
class EvenNumbersGen : System.Collections.IEnumerable {
    [int[]]$Nums = @(1, 2, 3, 4, 5, 6)
    [System.Collections.IEnumerator]GetEnumerator() {
        return $this.Nums.GetEnumerator()
    }
}
$eng = [EvenNumbersGen]::new()
$evens = @($eng | Where-Object { $_ % 2 -eq 0 })
if ($evens.Length -ne 3 -or $evens[0] -ne 2 -or $evens[2] -ne 6) {
    Write-Host "FAIL: Custom enumerator with Where-Object filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
