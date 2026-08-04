# vybe-test: powershell/splatting/dynamic_scriptblock_splat
function Invoke-Add {
    param($a, $b, $c)
    return $a + $b + $c
}
$script = { param($x, $y, $z) Invoke-Add @PSBoundParameters }
$a = 2; $b = 3; $c = 5
$result = & $script -x $a -y $b -z $c
if ($result -ne 10) {
    Write-Host "FAIL: expected 10, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
