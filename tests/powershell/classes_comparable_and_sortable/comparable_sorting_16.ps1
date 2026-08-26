# vybe-test: powershell/classes_comparable_and_sortable/comparable_sorting_16
class SortItem_16 : System.IComparable {
    [int]$Weight
    SortItem_16([int]$w) { $this.Weight = $w }
    [int]CompareTo([object]$other) {
        if ($other -isnot [SortItem_16]) { return 1 }
        return $this.Weight.CompareTo($other.Weight)
    }
}
$i1 = [SortItem_16]::new(30)
$i2 = [SortItem_16]::new(10)
$arr = [SortItem_16[]]@($i1, $i2)
[System.Array]::Sort($arr)
if ($arr[0].Weight -ne 10 -or $arr[1].Weight -ne 30) { Write-Host "FAIL: IComparable sort failed"; exit 1 }
Write-Host "PASS"; exit 0
