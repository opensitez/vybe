# vybe-test: powershell/hashtables/hashtable_add_key
$hash = @{ Name = "Charlie" }
$hash["Age"] = 40
if ($hash["Age"] -ne 40) {
    Write-Host "FAIL: expected 40, got $($hash['Age'])"
    exit 1
}
Write-Host "PASS"
exit 0
