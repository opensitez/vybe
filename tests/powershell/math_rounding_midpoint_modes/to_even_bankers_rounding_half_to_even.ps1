# vybe-test: powershell/math_rounding_midpoint_modes/to_even_bankers_rounding_half_to_even
$r1 = [math]::Round(2.5, [System.MidpointRounding]::ToEven) # rounds to 2
$r2 = [math]::Round(3.5, [System.MidpointRounding]::ToEven) # rounds to 4
if ($r1 -ne 2.0 -or $r2 -ne 4.0) {
    Write-Host "FAIL: ToEven banker's rounding failed, r1=$r1, r2=$r2"
    exit 1
}
Write-Host "PASS"
exit 0
