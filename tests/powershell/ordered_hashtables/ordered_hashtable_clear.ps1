# vybe-test: powershell/ordered_hashtables/ordered_hashtable_clear
$h = [ordered]@{ X = 1; Y = 2 }
$h.Clear()
if ($h.Count -ne 0) {
    Write-Host "FAIL: Count after Clear() expected 0, got $($h.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
