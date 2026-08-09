# vybe-test: powershell/ordered_hashtables/ordered_hashtable_clone
$h = [ordered]@{ A = 1; B = 2 }
$copy = $h.Clone()
$copy["A"] = 999
if ($h["A"] -ne 1) {
    Write-Host "FAIL: clone mutation affected original hashtable"
    exit 1
}
Write-Host "PASS"
exit 0
