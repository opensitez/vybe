# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_15
class CustomCollection_15 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(15, 16, 17)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_15]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 48
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
