# vybe-test: powershell/new_object_cmdlet/new_object_random_with_seed_produces_deterministic_sequence
# New-Object System.Random -ArgumentList <seed> generates a reproducible random sequence
$r1 = New-Object System.Random -ArgumentList 42
$r2 = New-Object System.Random -ArgumentList 42

$v1 = $r1.Next(0, 100)
$v2 = $r2.Next(0, 100)

if ($v1 -ne $v2) {
    Write-Host "FAIL: seeded Random should produce identical values, got $v1 vs $v2"
    exit 1
}

Write-Host "PASS"
exit 0
