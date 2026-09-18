# vybe-test: powershell/intrinsic_collection_overloads/where_preserves_underlying_object_identity
$targetObj = [pscustomobject]@{ Id = "target_42"; Label = "Special" }
$otherObj = [pscustomobject]@{ Id = "other_99"; Label = "Normal" }
$list = @($otherObj, $targetObj)

$filtered = $list.Where({ $_.Id -eq "target_42" })

if ($filtered.Count -ne 1) {
    Write-Host "FAIL: expected 1 match, got $($filtered.Count)"
    exit 1
}

# The returned element must be the exact same object reference, not a cloned or re-instantiated copy
$isSameRef = [object]::ReferenceEquals($targetObj, $filtered[0])
if (-not $isSameRef) {
    Write-Host "FAIL: .Where() returned a cloned copy instead of preserving original object reference"
    exit 1
}

# Mutating the original reflects in the filtered result
$targetObj.Label = "Mutated"
if ($filtered[0].Label -ne "Mutated") {
    Write-Host "FAIL: mutation on original object not reflected in filtered reference"
    exit 1
}

Write-Host "PASS"
exit 0
