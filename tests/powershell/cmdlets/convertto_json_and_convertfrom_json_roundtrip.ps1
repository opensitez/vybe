# vybe-test: powershell/cmdlets/convertto_json_and_convertfrom_json_roundtrip
# ConvertTo-Json serializes an object tree and ConvertFrom-Json restores the property graph
$original = [PSCustomObject]@{
    ClusterName = "AlphaCluster"
    NodeCount   = 8
    Enabled     = $true
    Tags        = @("primary", "us-east")
}

$jsonText = $original | ConvertTo-Json -Depth 3 -Compress
$restored = $jsonText | ConvertFrom-Json

if ($restored.ClusterName -ne "AlphaCluster") {
    Write-Host "FAIL: ClusterName mismatch, got '$($restored.ClusterName)'"
    exit 1
}

if ($restored.NodeCount -ne 8) {
    Write-Host "FAIL: NodeCount mismatch, got $($restored.NodeCount)"
    exit 1
}

if ($restored.Enabled -ne $true) {
    Write-Host "FAIL: Enabled boolean mismatch"
    exit 1
}

if ($restored.Tags.Count -ne 2 -or $restored.Tags[0] -ne "primary") {
    Write-Host "FAIL: Tags array mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
