# vybe-test: powershell/splatting/general_parameter_splatting
function Sum-Values {
    param($a, $b, $c)
    return $a + $b + $c
}
$args = @{ a = 10; b = 20; c = 30 }
$result = Sum-Values @args
if ($result -ne 60) {
    Write-Host "FAIL: expected 60, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
