# vybe-test: powershell/ordered_hashtables/ordered_hashtable_remove
$h = [ordered]@{ A = 1; B = 2; C = 3 }
$h.Remove("B")
$keys = @($h.Keys)
if ($keys.Count -ne 2 -or $keys[0] -ne "A" -or $keys[1] -ne "C") {
    Write-Host "FAIL: key removal expected A, C remaining, got $($keys -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
