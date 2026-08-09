# vybe-test: powershell/ordered_hashtables/ordered_hashtable_null_value
$h = [ordered]@{ EmptyKey = $null }
if (-not $h.Contains("EmptyKey")) {
    Write-Host "FAIL: key with null value expected to exist in ordered hashtable"
    exit 1
}
if ($h["EmptyKey"] -ne $null) {
    Write-Host "FAIL: expected null value"
    exit 1
}
Write-Host "PASS"
exit 0
