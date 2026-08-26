# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_5
class CustomCollection_5 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(5, 6, 7)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_5]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 18
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
