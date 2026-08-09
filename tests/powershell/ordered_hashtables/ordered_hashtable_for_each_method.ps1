# vybe-test: powershell/ordered_hashtables/ordered_hashtable_for_each_method
$h = [ordered]@{ A = 10; B = 20 }
$res = $h.Keys.ForEach({ $_.ToLower() })
if ($res[0] -ne "a" -or $res[1] -ne "b") {
    Write-Host "FAIL: .ForEach() on ordered keys expected 'a','b', got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
