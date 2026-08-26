# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_11
class CustomCollection_11 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(11, 12, 13)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_11]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 36
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
