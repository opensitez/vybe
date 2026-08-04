# vybe-test: powershell/hashtables/hashtable_remove_key
$hash = @{ Name = "Alice"; Age = 30; City = "NYC" }
$hash.Remove("Age")
$count = $hash.Count
if ($count -ne 2) {
    Write-Host "FAIL: expected 2, got $count"
    exit 1
}
if ($hash.ContainsKey("Age")) {
    Write-Host "FAIL: Age key should be removed"
    exit 1
}
Write-Host "PASS"
exit 0
