# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_13
class CustomCollection_13 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(13, 14, 15)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_13]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 42
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
