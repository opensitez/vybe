# vybe-test: powershell/scriptblocks/scriptblock_getnewinvoker
$sb = [scriptblock]::Create('param($a,$b) $a + $b')
$result = & $sb 10 32
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
