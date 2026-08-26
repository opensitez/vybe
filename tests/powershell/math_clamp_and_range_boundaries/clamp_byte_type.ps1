# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_byte_type
[byte]$b = 200
$val = [math]::Clamp($b, [byte]10, [byte]100)
if ($val -ne 100) {
    Write-Host "FAIL: Clamp byte failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
