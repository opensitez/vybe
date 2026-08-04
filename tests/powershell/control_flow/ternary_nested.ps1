# vybe-test: powershell/control_flow/ternary_nested
$score = 72
$grade = ($score -ge 90) ? "A" : ($score -ge 80) ? "B" : ($score -ge 70) ? "C" : "F"
if ($grade -ne "C") {
    Write-Host "FAIL: expected 'C', got '$grade'"
    exit 1
}
$score2 = 95
$grade2 = ($score2 -ge 90) ? "A" : "not A"
if ($grade2 -ne "A") {
    Write-Host "FAIL: expected 'A', got '$grade2'"
    exit 1
}
Write-Host "PASS"
exit 0
