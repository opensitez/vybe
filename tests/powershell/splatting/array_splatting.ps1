# vybe-test: powershell/splatting/array_splatting
function Add-Numbers {
    param($a, $b, $c)
    return $a + $b + $c
}
$args = @{ a = 1; b = 2; c = 3 }
$result = Add-Numbers @args
if ($result -ne 6) {
    Write-Host "FAIL: expected 6, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
