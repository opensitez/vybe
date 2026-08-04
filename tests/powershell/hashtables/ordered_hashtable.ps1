# vybe-test: powershell/hashtables/ordered_hashtable
$hash = [ordered]@{ First = 1; Second = 2; Third = 3 }
$keys = $hash.Keys
if ($keys.Count -ne 3) {
    Write-Host "FAIL: expected 3 keys"
    exit 1
}
Write-Host "PASS"
exit 0
