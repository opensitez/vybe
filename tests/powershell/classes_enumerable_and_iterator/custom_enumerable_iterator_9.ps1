# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_9
class CustomCollection_9 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(9, 10, 11)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_9]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 30
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
