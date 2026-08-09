# vybe-test: powershell/ordered_hashtables/ordered_hashtable_type_check
$h = [ordered]@{ a = 1 }
if (-not ($h -is [System.Collections.Specialized.IOrderedDictionary])) {
    Write-Host "FAIL: object does not implement IOrderedDictionary interface"
    exit 1
}
Write-Host "PASS"
exit 0
