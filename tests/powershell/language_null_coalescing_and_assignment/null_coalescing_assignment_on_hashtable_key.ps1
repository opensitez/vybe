# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_assignment_on_hashtable_key
$ht = @{ existing = "A"; unset = $null }
$ht["unset"] ??= "DefaultUnset"
$ht["existing"] ??= "NewA"
if ($ht["unset"] -ne "DefaultUnset" -or $ht["existing"] -ne "A") {
    Write-Host "FAIL: ??= on hashtable key failed"
    exit 1
}
Write-Host "PASS"
exit 0
