# vybe-test: powershell/classes_comparable_and_sortable/comparable_sorting_19
class SortItem_19 : System.IComparable {
    [int]$Weight
    SortItem_19([int]$w) { $this.Weight = $w }
    [int]CompareTo([object]$other) {
        if ($other -isnot [SortItem_19]) { return 1 }
        return $this.Weight.CompareTo($other.Weight)
    }
}
$i1 = [SortItem_19]::new(30)
$i2 = [SortItem_19]::new(10)
$arr = [SortItem_19[]]@($i1, $i2)
[System.Array]::Sort($arr)
if ($arr[0].Weight -ne 10 -or $arr[1].Weight -ne 30) { Write-Host "FAIL: IComparable sort failed"; exit 1 }
Write-Host "PASS"; exit 0
