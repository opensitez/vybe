# vybe-test: powershell/hashtables/create_hashtable
$hash = @{ Name = "John"; Age = 30 }
$count = $hash.Count
if ($count -ne 2) {
    Write-Host "FAIL: expected 2, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
