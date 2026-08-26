# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_20
class CustomCollection_20 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(20, 21, 22)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_20]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 63
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
