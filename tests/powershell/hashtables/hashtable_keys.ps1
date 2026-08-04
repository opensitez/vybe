# vybe-test: powershell/hashtables/hashtable_keys
$hash = @{ A = 1; B = 2; C = 3 }
$keys = $hash.Keys
$count = $keys.Count
if ($count -ne 3) {
    Write-Host "FAIL: expected 3 keys, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
