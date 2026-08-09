# vybe-test: powershell/ordered_hashtables/ordered_hashtable_where_method
$h = [ordered]@{ X = 5; Y = 15; Z = 25 }
$filtered = $h.GetEnumerator().Where({ $_.Value -gt 10 })
if ($filtered.Count -ne 2) {
    Write-Host "FAIL: .Where() expected 2 items, got $($filtered.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
