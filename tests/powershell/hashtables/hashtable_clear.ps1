# vybe-test: powershell/hashtables/hashtable_clear
$hash = @{ A = 1; B = 2; C = 3 }
$hash.Clear()
if ($hash.Count -ne 0) {
    Write-Host "FAIL: expected 0 keys after clear, got $($hash.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
