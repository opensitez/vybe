# vybe-test: powershell/control_flow/nested_loops
$count = 0
for ($i = 0; $i -lt 3; $i++) {
    for ($j = 0; $j -lt 2; $j++) {
        $count++
    }
}
if ($count -ne 6) {
    Write-Host "FAIL: expected 6, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
