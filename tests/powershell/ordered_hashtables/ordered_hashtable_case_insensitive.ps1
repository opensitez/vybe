# vybe-test: powershell/ordered_hashtables/ordered_hashtable_case_insensitive
$h = [ordered]@{ MyKey = "Value" }
if ($h["mykey"] -ne "Value") {
    Write-Host "FAIL: case-insensitive key access expected Value, got $($h['mykey'])"
    exit 1
}
Write-Host "PASS"
exit 0
