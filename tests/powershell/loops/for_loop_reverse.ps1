# vybe-test: powershell/loops/for_loop_reverse
$result = @()
for ($i = 5; $i -ge 1; $i--) {
    $result += $i
}
$expected = @(5,4,3,2,1)
for ($i = 0; $i -lt 5; $i++) {
    if ($result[$i] -ne $expected[$i]) {
        Write-Host "FAIL: [$i] got $($result[$i])"
        exit 1
    }
}
Write-Host "PASS"
exit 0
