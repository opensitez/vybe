# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_17
class CustomCollection_17 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(17, 18, 19)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_17]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 54
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
