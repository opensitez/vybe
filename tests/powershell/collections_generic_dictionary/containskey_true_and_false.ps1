# vybe-test: powershell/collections_generic_dictionary/containskey_true_and_false
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("exists", 42)
if (-not $d.ContainsKey("exists") -or $d.ContainsKey("missing")) {
    Write-Host "FAIL: ContainsKey failed"
    exit 1
}
Write-Host "PASS"
exit 0
