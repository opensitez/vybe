# vybe-test: powershell/splatting/splat_empty_hash_injection
function Combine-Values {
    param($x, $y, $z = 5)
    return $x + $y + $z
}
$params = @{ x = 1; y = 2 }
$result = Combine-Values @params
if ($result -ne 8) {
    Write-Host "FAIL: expected 8, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
