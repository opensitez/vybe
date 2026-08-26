# vybe-test: powershell/collections_sorted_dictionary/missing_key_exception
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$caught = $false
try {
    $x = $sd.get_Item("not_found")
} catch [System.Collections.Generic.KeyNotFoundException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected KeyNotFoundException"
    exit 1
}
Write-Host "PASS"
exit 0
