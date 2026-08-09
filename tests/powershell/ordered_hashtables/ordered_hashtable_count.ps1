# vybe-test: powershell/ordered_hashtables/ordered_hashtable_count
$h = [ordered]@{ A = 1; B = 2; C = 3; D = 4 }
if ($h.Count -ne 4) {
    Write-Host "FAIL: Count expected 4, got $($h.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
