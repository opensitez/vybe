# vybe-test: powershell/hashtables/hashtable_addition_operator_combines_disjoint_keys
# The + operator merges two hashtables with disjoint keys into a new composite hashtable
$h1 = @{ FirstName = "Grace"; Age = 32 }
$h2 = @{ LastName = "Hopper"; Role = "Pioneer" }

$combined = $h1 + $h2

if ($combined.Count -ne 4) {
    Write-Host "FAIL: expected combined count 4, got $($combined.Count)"
    exit 1
}

if ($combined.FirstName -ne "Grace" -or $combined.LastName -ne "Hopper") {
    Write-Host "FAIL: combined hashtable values mismatch"
    exit 1
}

# Original hashtables should remain untouched
if ($h1.Count -ne 2 -or $h2.Count -ne 2) {
    Write-Host "FAIL: operand hashtables were mutated by addition"
    exit 1
}

Write-Host "PASS"
exit 0
