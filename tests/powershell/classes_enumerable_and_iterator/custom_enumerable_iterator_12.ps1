# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_12
class CustomCollection_12 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(12, 13, 14)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_12]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 39
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
