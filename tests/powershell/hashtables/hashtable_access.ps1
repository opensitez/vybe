# vybe-test: powershell/hashtables/hashtable_access
$hash = @{ Name = "Alice"; Age = 25 }
$name = $hash["Name"]
if ($name -ne "Alice") {
    Write-Host "FAIL: expected 'Alice', got '$name'"
    exit 1
}
Write-Host "PASS"
exit 0
