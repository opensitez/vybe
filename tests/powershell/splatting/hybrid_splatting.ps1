# vybe-test: powershell/splatting/hybrid_splatting
function Add-Values {
    param($x, $y, $z)
    return $x + $y + $z
}
$arrayArgs = 1, 2
$hashArgs = @{ z = 3 }
$result = Add-Values @arrayArgs @hashArgs
if ($result -ne 6) {
    Write-Host "FAIL: expected 6, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
