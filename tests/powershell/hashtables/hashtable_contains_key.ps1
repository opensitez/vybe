# vybe-test: powershell/hashtables/hashtable_contains_key
$hash = @{ Name = "David"; Age = 28 }
$result = $hash.ContainsKey("Name")
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
