# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_empty_hashtable_is_not_null
$ht = @{}
# Empty hashtable is NOT null, ?? must return empty hashtable
$res = $ht ?? @{ a = 1 }
if ($res.Count -ne 0) {
    Write-Host "FAIL: Empty hashtable should not be coalesced, got count $($res.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
