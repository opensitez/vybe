# vybe-test: powershell/classes_enumerable_and_iterator/custom_enumerable_iterator_18
class CustomCollection_18 : System.Collections.IEnumerable {
    [int[]]$Items = [int[]]@(18, 19, 20)
    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Items.GetEnumerator()
    }
}
$col = [CustomCollection_18]::new()
$sum = 0
foreach ($item in $col.GetEnumerator()) { $sum += $item }
$expected = 57
if ($sum -ne $expected) { Write-Host "FAIL: Custom enumerable iterator failed, expected $expected, got $sum"; exit 1 }
Write-Host "PASS"; exit 0
