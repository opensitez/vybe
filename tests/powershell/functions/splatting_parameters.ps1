# vybe-test: powershell/functions/splatting_parameters
function Add-Three {
    param($a, $b, $c)
    return $a + $b + $c
}
$params = @{ a = 10; b = 20; c = 30 }
$result = Add-Three @params
if ($result -ne 60) {
    Write-Host "FAIL: expected 60, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
