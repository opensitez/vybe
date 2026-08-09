# vybe-test: powershell/ordered_hashtables/ordered_hashtable_add_member
$h = [ordered]@{ One = 1 }
$h.Add("Two", 2)
$keys = @($h.Keys)
if ($keys[1] -ne "Two") {
    Write-Host "FAIL: dynamic Add key order expected Two, got $($keys[1])"
    exit 1
}
Write-Host "PASS"
exit 0
