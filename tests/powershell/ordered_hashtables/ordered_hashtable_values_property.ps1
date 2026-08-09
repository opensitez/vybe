# vybe-test: powershell/ordered_hashtables/ordered_hashtable_values_property
$h = [ordered]@{ First = 100; Second = 200 }
$v = @($h.Values)
if ($v[0] -ne 100 -or $v[1] -ne 200) {
    Write-Host "FAIL: Values property expected 100, 200, got $($v -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
