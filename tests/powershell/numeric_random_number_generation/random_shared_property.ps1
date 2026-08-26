# vybe-test: powershell/numeric_random_number_generation/random_shared_property
$shared = [System.Random]::Shared
$v = $shared.Next(1, 100)
if ($v -lt 1 -or $v -ge 100) {
    Write-Host "FAIL: Random.Shared failed"
    exit 1
}
Write-Host "PASS"
exit 0
