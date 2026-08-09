# vybe-test: powershell/type_accelerators/type_accelerator_hashtable
$h = [hashtable]@{ a = 10; b = 20 }
if ($h["a"] -ne 10) {
    Write-Host "FAIL: hashtable 'a' expected 10, got $($h['a'])"
    exit 1
}
if ($h.Count -ne 2) {
    Write-Host "FAIL: hashtable Count expected 2, got $($h.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
