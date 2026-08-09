# vybe-test: powershell/null_conditional/null_conditional_nested_null_object
$node = @{ Child = @{ Leaf = "Green" } }
$res = ${node}?["Child"]?["Leaf"]
if ($res -ne "Green") {
    Write-Host "FAIL: nested hashtable null-conditional index expected 'Green', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
