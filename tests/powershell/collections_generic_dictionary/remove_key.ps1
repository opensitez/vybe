# vybe-test: powershell/collections_generic_dictionary/remove_key
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("a", 1); $d.Add("b", 2)
$rem = $d.Remove("a")
if (-not $rem -or $d.Count -ne 1 -or $d.ContainsKey("a")) {
    Write-Host "FAIL: Remove key failed"
    exit 1
}
Write-Host "PASS"
exit 0
