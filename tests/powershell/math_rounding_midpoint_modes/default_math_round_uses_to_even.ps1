# vybe-test: powershell/math_rounding_midpoint_modes/default_math_round_uses_to_even
$r = [math]::Round(4.5)
if ($r -ne 4.0) {
    Write-Host "FAIL: Default [math]::Round should use ToEven (4.5 -> 4.0), got $r"
    exit 1
}
Write-Host "PASS"
exit 0
