# vybe-test: powershell/collections_generic_dictionary/trygetvalue_present_and_missing
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("score", 100)
$hasKey = $d.ContainsKey("score")
$val = if ($hasKey) { $d["score"] } else { 0 }
if (-not $hasKey -or $val -ne 100 -or $d.ContainsKey("unknown")) {
    Write-Host "FAIL: Dictionary key lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
