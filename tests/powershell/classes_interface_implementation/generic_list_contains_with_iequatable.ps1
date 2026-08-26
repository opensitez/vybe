# vybe-test: powershell/classes_interface_implementation/generic_list_contains_with_iequatable
class EquatableItemDemo {
    [string]$Code
    EquatableItemDemo([string]$c) { $this.Code = $c }
    [bool]Equals([object]$other) {
        if ($other -eq $null -or $other -isnot [EquatableItemDemo]) { return $false }
        return $this.Code -eq $other.Code
    }
}
$list = [System.Collections.Generic.List[EquatableItemDemo]]::new()
$list.Add([EquatableItemDemo]::new("ABC"))
if ($list[0].Code -ne "ABC") {
    Write-Host "FAIL: List with custom class item failed"
    exit 1
}
Write-Host "PASS"
exit 0
