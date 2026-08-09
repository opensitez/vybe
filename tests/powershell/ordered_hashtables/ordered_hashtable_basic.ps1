# vybe-test: powershell/ordered_hashtables/ordered_hashtable_basic
$h = [ordered]@{ Z = 26; A = 1; M = 13 }
$keys = @($h.Keys)
if ($keys[0] -ne "Z" -or $keys[1] -ne "A" -or $keys[2] -ne "M") {
    Write-Host "FAIL: key insertion order expected Z, A, M, got $($keys -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
