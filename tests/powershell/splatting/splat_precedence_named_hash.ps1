# vybe-test: powershell/splatting/splat_precedence_named_hash
function Evaluate {
    param($x, $y, $z)
    return "${x}-${y}-${z}"
}
$base = @{ x = 'A'; y = 'B'; z = 'C' }
$result = Evaluate -y 'Override' @base
if ($result -ne 'A-Override-C') {
    Write-Host "FAIL: expected A-Override-C, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
