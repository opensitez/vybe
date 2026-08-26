# vybe-test: powershell/json_custom_object_roundtrip/custom_object_roundtrip_pipeline_composition
$orig = @(
    [pscustomobject]@{ Score = 80; Name = "A" },
    [pscustomobject]@{ Score = 40; Name = "B" }
)
$filtered = $orig | ConvertTo-Json | ConvertFrom-Json | Where-Object { $_.Score -ge 50 }
if ($filtered.Name -ne "A") {
    Write-Host "FAIL: Custom object pipeline composition failed"
    exit 1
}
Write-Host "PASS"
exit 0
