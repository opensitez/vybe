# vybe-test: powershell/numeric_random_number_generation/random_distribution_variety_check
$rnd = [System.Random]::new()
$seen = [System.Collections.Generic.HashSet[int]]::new()
for ($i = 0; $i -lt 100; $i++) {
    $null = $seen.Add($rnd.Next(1, 1000))
}
if ($seen.Count -lt 80) {
    Write-Host "FAIL: Random distribution generated too many collisions"
    exit 1
}
Write-Host "PASS"
exit 0
