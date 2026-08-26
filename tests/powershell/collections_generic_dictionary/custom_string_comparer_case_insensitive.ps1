# vybe-test: powershell/collections_generic_dictionary/custom_string_comparer_case_insensitive
$comparer = [System.StringComparer]::OrdinalIgnoreCase
$d = [System.Collections.Generic.Dictionary[string, int]]::new($comparer)
$d.Add("HeLLo", 99)
if ($d["hello"] -ne 99 -or $d["HELLO"] -ne 99) {
    Write-Host "FAIL: Case-insensitive comparer in Dictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
