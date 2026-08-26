# vybe-test: powershell/classes_comparable_and_sortable/comparable_sorting_1
class SortItem_1 : System.IComparable {
    [int]$Weight
    SortItem_1([int]$w) { $this.Weight = $w }
    [int]CompareTo([object]$other) {
        if ($other -isnot [SortItem_1]) { return 1 }
        return $this.Weight.CompareTo($other.Weight)
    }
}
$i1 = [SortItem_1]::new(30)
$i2 = [SortItem_1]::new(10)
$arr = [SortItem_1[]]@($i1, $i2)
[System.Array]::Sort($arr)
if ($arr[0].Weight -ne 10 -or $arr[1].Weight -ne 30) { Write-Host "FAIL: IComparable sort failed"; exit 1 }
Write-Host "PASS"; exit 0
