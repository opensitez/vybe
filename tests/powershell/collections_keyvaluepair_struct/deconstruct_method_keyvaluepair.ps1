# vybe-test: powershell/collections_keyvaluepair_struct/deconstruct_method_keyvaluepair
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("pi", 3)
$k = ""
$v = 0
$kvp.Deconstruct([ref]$k, [ref]$v)
if ($k -ne "pi" -or $v -ne 3) {
    Write-Host "FAIL: Deconstruct method failed"
    exit 1
}
Write-Host "PASS"
exit 0
