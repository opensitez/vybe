# vybe-test: powershell/collections_keyvaluepair_struct/keyvaluepair_array_iteration
$k1 = [System.Collections.Generic.KeyValuePair[string, int]]::new("x", 10)
$k2 = [System.Collections.Generic.KeyValuePair[string, int]]::new("y", 20)
$arr = @($k1, $k2)
$sum = 0
foreach ($item in $arr) { $sum += $item.Value }
if ($sum -ne 30 -or $arr[0].Key -ne "x") {
    Write-Host "FAIL: KeyValuePair array iteration failed"
    exit 1
}
Write-Host "PASS"
exit 0
