# vybe-test: powershell/classes/class_list_of_instances
class Item {
    [string]$Name
    [int]$Value
    Item([string]$n, [int]$v) { $this.Name = $n; $this.Value = $v }
}
$items = @(
    [Item]::new("apple", 3),
    [Item]::new("banana", 7),
    [Item]::new("cherry", 2)
)
$total = ($items | Measure-Object -Property Value -Sum).Sum
if ($total -ne 12) {
    Write-Host "FAIL: expected 12, got $total"
    exit 1
}
$sorted = $items | Sort-Object Value
if ($sorted[0].Name -ne "cherry") { Write-Host "FAIL: sort"; exit 1 }
Write-Host "PASS"
exit 0
