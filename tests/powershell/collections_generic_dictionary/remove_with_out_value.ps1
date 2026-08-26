# vybe-test: powershell/collections_generic_dictionary/remove_with_out_value
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("score", 75)
$val = 0
$rem = $d.Remove("score", [ref]$val)
if (-not $rem -or $val -ne 75 -or $d.Count -ne 0) {
    Write-Host "FAIL: Remove with out value failed"
    exit 1
}
Write-Host "PASS"
exit 0
