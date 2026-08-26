# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_in_foreach_statement
class StatementGen : System.Collections.IEnumerable {
    [int[]]$Values = @(10, 20)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Values.GetEnumerator() }
}
$sg = [StatementGen]::new()
$sum = 0
foreach ($item in $sg) {
    $sum += $item
}
if ($sum -ne 30) {
    Write-Host "FAIL: foreach statement over custom IEnumerable failed, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
