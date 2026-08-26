# vybe-test: powershell/collections_generic_dictionary/foreach_keyvaluepair_iteration
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("a", 1); $d.Add("b", 2)
$keys = @($d.Keys)
if ($keys.Count -ne 2 -or -not ($keys -contains "a") -or -not ($keys -contains "b")) {
    Write-Host "FAIL: Keys iteration failed"
    exit 1
}
Write-Host "PASS"
exit 0
