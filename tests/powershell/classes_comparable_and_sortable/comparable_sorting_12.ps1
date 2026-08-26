# vybe-test: powershell/classes_comparable_and_sortable/comparable_sorting_12
class SortItem_12 : System.IComparable {
    [int]$Weight
    SortItem_12([int]$w) { $this.Weight = $w }
    [int]CompareTo([object]$other) {
        if ($other -isnot [SortItem_12]) { return 1 }
        return $this.Weight.CompareTo($other.Weight)
    }
}
$i1 = [SortItem_12]::new(30)
$i2 = [SortItem_12]::new(10)
$arr = [SortItem_12[]]@($i1, $i2)
[System.Array]::Sort($arr)
if ($arr[0].Weight -ne 10 -or $arr[1].Weight -ne 30) { Write-Host "FAIL: IComparable sort failed"; exit 1 }
Write-Host "PASS"; exit 0
