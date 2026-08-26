# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_custom_objects
$objs = @([pscustomobject]@{ Id = 1 }, [pscustomobject]@{ Id = 2 })
$side = $null
$out = @($objs | Tee-Object -Variable side | ForEach-Object { $_.Id * 2 })
if ($out[0] -ne 2 -or $side.Count -ne 2 -or $side[0].Id -ne 1) {
    Write-Host "FAIL: Tee-Object with custom objects failed"
    exit 1
}
Write-Host "PASS"
exit 0
