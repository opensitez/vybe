# vybe-test: powershell/generic_types/generic_dictionary_key_lookup
$dict = [System.Collections.Generic.Dictionary[int, string]]::new()
$dict[100] = "Centum"
if (-not $dict.ContainsKey(100)) {
    Write-Host "FAIL: ContainsKey(100) expected true"
    exit 1
}
if ($dict[100] -ne "Centum") {
    Write-Host "FAIL: dict[100] expected Centum, got $($dict[100])"
    exit 1
}
Write-Host "PASS"
exit 0
