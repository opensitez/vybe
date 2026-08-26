# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_1
class CustomCollection_1 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(1, 2, 3)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_1]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 6
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
