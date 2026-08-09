# vybe-test: powershell/ordered_hashtables/ordered_hashtable_keys_property
$h = [ordered]@{ Beta = 2; Alpha = 1 }
$k = $h.Keys
if ($k[0] -ne "Beta" -or $k[1] -ne "Alpha") {
    Write-Host "FAIL: Keys property order expected Beta, Alpha, got $($k -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
