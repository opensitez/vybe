# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_16
class CustomCollection_16 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(16, 17, 18)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_16]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 51
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
