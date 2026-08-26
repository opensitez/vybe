# vybe-test: powershell/scriptblock_closures/closure_array_capture
$items = @("Alpha", "Beta")
$sb = { $items[1] }.GetNewClosure()
$res = &$sb
if ($res -ne "Beta") {
    Write-Host "FAIL: array capture in closure expected 'Beta', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
