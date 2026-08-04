# vybe-test: powershell/scriptblocks/scriptblock_invoke
$sb = { param($x) $x * 2 }
$result = & $sb 21
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
