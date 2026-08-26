# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_6
class CustomCollection_6 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(6, 7, 8)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_6]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 21
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
