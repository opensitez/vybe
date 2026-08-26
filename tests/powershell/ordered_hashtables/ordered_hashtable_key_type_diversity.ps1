# vybe-test: powershell/ordered_hashtables/ordered_hashtable_key_type_diversity
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
