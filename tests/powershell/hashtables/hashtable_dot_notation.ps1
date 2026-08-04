# vybe-test: powershell/hashtables/hashtable_dot_notation
$hash = @{ Name = "Bob"; Age = 35 }
$age = $hash.Age
if ($age -ne 35) {
    Write-Host "FAIL: expected 35, got $age"
    exit 1
}
Write-Host "PASS"
exit 0
