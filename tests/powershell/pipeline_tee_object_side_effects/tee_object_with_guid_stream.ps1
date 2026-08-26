# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_guid_stream
$g1 = [guid]::NewGuid()
$g2 = [guid]::NewGuid()
$sideGuids = $null
$res = @($g1, $g2 | Tee-Object -Variable sideGuids | ForEach-Object { $_.ToString() })
if ($res.Length -ne 2 -or $sideGuids[0] -ne $g1) {
    Write-Host "FAIL: Tee-Object with GUID stream failed"
    exit 1
}
Write-Host "PASS"
exit 0
