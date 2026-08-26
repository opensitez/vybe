# vybe-test: powershell/collections_sorted_dictionary/duplicate_key_exception
$sd = [System.Collections.Generic.SortedDictionary[string, int]]::new()
$sd.Add("k", 1)
$caught = $false
try {
    $sd.Add("k", 2)
} catch [System.ArgumentException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentException on duplicate key"
    exit 1
}
Write-Host "PASS"
exit 0
