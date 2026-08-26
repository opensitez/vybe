# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_4
class CustomCollection_4 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(4, 5, 6)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_4]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 15
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
