# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_2
class CustomCollection_2 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(2, 3, 4)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_2]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 9
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
