# vybe-test: powershell/ordered_hashtables/ordered_hashtable_key_type_diversity
$h = [ordered]@{ 1 = "IntKey"; "2" = "StringKey" }
if ($h[1] -ne "IntKey") {
    Write-Host "FAIL: key 1 expected IntKey, got $($h[1])"
    exit 1
}
Write-Host "PASS"
exit 0
