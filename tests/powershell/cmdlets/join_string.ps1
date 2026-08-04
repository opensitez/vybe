# vybe-test: powershell/cmdlets/join_string
$items = "a", "b", "c"
$result = $items -join ","
if ($result -ne "a,b,c") {
    Write-Host "FAIL: expected 'a,b,c', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
