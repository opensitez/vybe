# vybe-test: powershell/json_custom_object_roundtrip/array_of_custom_objects_roundtrip
$orig = @(
    [pscustomobject]@{ Id = 1; Tag = "A" },
    [pscustomobject]@{ Id = 2; Tag = "B" }
)
$json = $orig | ConvertTo-Json
$recovered = @($json | ConvertFrom-Json)
if ($recovered.Length -ne 2 -or $recovered[0].Tag -ne "A" -or $recovered[1].Id -ne 2) {
    Write-Host "FAIL: Array of custom objects roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
