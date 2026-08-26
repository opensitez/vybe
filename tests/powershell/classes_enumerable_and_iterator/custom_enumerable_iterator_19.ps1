# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_19
class CustomCollection_19 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(19, 20, 21)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_19]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 60
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
