# vybe-test: powershell/types/type_hashtable
$hash = @{ Key = "Value" }
$result = $hash -is [hashtable]
if ($result -ne $true) {
    Write-Host "FAIL: expected True for hashtable type check"
    exit 1
}
Write-Host "PASS"
exit 0
