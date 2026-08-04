# vybe-test: powershell/math/math_round_modes
# PowerShell uses banker's rounding by default (MidpointRounding.ToEven)
$r1 = [Math]::Round(2.5)   # rounds to 2 (even)
$r2 = [Math]::Round(3.5)   # rounds to 4 (even)
$r3 = [Math]::Round(2.5, 0, [MidpointRounding]::AwayFromZero)  # rounds to 3
if ($r1 -ne 2) { Write-Host "FAIL: 2.5 banker round = $r1"; exit 1 }
if ($r2 -ne 4) { Write-Host "FAIL: 3.5 banker round = $r2"; exit 1 }
if ($r3 -ne 3) { Write-Host "FAIL: 2.5 away-from-zero = $r3"; exit 1 }
Write-Host "PASS"
exit 0
