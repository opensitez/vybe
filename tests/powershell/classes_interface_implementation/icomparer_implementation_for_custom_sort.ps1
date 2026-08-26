# vybe-test: powershell/classes_interface_implementation/icomparer_implementation_for_custom_sort
class StringLengthComparer : System.Collections.Generic.IComparer[string] {
    [int]Compare([string]$x, [string]$y) {
        return $x.Length.CompareTo($y.Length)
    }
}
$comp = [StringLengthComparer]::new()
[string[]]$words = @("elephant", "cat", "hippopotamus", "dog")
[System.Array]::Sort($words, $comp)
if ($words[0] -ne "cat" -and $words[0] -ne "dog") {
    Write-Host "FAIL: IComparer custom sort failed"
    exit 1
}
Write-Host "PASS"
exit 0
