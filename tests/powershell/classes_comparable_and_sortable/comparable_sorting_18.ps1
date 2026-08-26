# vybe-test: powershell/classes_comparable_and_sortable/comparable_sorting_18
class SortItem_18 : System.IComparable {
    [int]$Weight
    SortItem_18([int]$w) { $this.Weight = $w }
    [int]CompareTo([object]$other) {
        if ($other -isnot [SortItem_18]) { return 1 }
        return $this.Weight.CompareTo($other.Weight)
    }
}
$i1 = [SortItem_18]::new(30)
$i2 = [SortItem_18]::new(10)
$arr = [SortItem_18[]]@($i1, $i2)
[System.Array]::Sort($arr)
if ($arr[0].Weight -ne 10 -or $arr[1].Weight -ne 30) { Write-Host "FAIL: IComparable sort failed"; exit 1 }
Write-Host "PASS"; exit 0
