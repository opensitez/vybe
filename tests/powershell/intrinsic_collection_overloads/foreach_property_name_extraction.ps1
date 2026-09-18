# vybe-test: powershell/intrinsic_collection_overloads/foreach_property_name_extraction
$items = @(
    [pscustomobject]@{ Id = 101; Code = "NYC"; Active = $true },
    [pscustomobject]@{ Id = 102; Code = "LON"; Active = $false },
    [pscustomobject]@{ Id = 103; Code = "TOK"; Active = $true }
)

# .ForEach('PropertyName') extracts member values directly without scriptblock overhead
$codes = $items.ForEach('Code')

if ($null -eq $codes -or $codes.Count -ne 3) {
    Write-Host "FAIL: .ForEach('Code') expected 3 items, got $($codes.Count)"
    exit 1
}

if ($codes[0] -ne "NYC" -or $codes[1] -ne "LON" -or $codes[2] -ne "TOK") {
    Write-Host "FAIL: expected @('NYC', 'LON', 'TOK'), got @($($codes -join ', '))"
    exit 1
}

# Also verify boolean property extraction
$actives = $items.ForEach('Active')
if ($actives[0] -ne $true -or $actives[1] -ne $false -or $actives[2] -ne $true) {
    Write-Host "FAIL: extracted boolean properties mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
