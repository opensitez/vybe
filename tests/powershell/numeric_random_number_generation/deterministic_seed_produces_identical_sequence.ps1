# vybe-test: powershell/numeric_random_number_generation/deterministic_seed_produces_identical_sequence
$r1 = [System.Random]::new(12345)
$r2 = [System.Random]::new(12345)
$mismatch = $false
for ($i = 0; $i -lt 10; $i++) {
    if ($r1.Next() -ne $r2.Next()) { $mismatch = $true; break }
}
if ($mismatch) {
    Write-Host "FAIL: Seeded Random must produce identical sequences"
    exit 1
}
Write-Host "PASS"
exit 0
