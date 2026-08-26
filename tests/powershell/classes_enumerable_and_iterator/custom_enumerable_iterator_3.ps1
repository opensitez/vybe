# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_3
class CustomCollection_3 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(3, 4, 5)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_3]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 12
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
