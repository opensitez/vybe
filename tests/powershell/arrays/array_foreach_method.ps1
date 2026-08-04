# vybe-test: powershell/arrays/array_foreach_method
$results = @()
(1..5).ForEach({ $results += $_ * $_ })
$expected = @(1,4,9,16,25)
for ($i = 0; $i -lt 5; $i++) {
    if ($results[$i] -ne $expected[$i]) {
        Write-Host "FAIL: [$i] got $($results[$i])"
        exit 1
    }
}
Write-Host "PASS"
exit 0
