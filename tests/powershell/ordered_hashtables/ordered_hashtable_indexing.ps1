# vybe-test: powershell/ordered_hashtables/ordered_hashtable_indexing
$h = [ordered]@{ First = "A"; Second = "B" }
if ($h[0] -ne "A") {
    Write-Host "FAIL: integer index [0] expected 'A', got $($h[0])"
    exit 1
}
if ($h[1] -ne "B") {
    Write-Host "FAIL: integer index [1] expected 'B', got $($h[1])"
    exit 1
}
Write-Host "PASS"
exit 0
