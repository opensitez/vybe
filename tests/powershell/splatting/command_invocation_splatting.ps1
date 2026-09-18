# vybe-test: powershell/splatting/command_invocation_splatting
function Multiply {
    param($x, $y)
    return $x * $y
}
$params = @{ x = 7; y = 8 }
$result = Multiply @params
if ($result -ne 56) {
    Write-Host "FAIL: expected 56, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
