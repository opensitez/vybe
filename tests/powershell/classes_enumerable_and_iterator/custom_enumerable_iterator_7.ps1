# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_7
class CustomCollection_7 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(7, 8, 9)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_7]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 24
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
