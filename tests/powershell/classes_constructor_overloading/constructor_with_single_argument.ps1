# vybe-test: powershell/classes_constructor_overloading/constructor_with_single_argument
class Item {
    [string]$Code
    Item([string]$c) { $this.Code = $c }
}
$item = [Item]::new("A100")
if ($item.Code -ne "A100") {
    Write-Host "FAIL: Single argument constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
