# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_10
class CustomCollection_10 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(10, 11, 12)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_10]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 33
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
