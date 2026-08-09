# vybe-test: powershell/pipeline_object_capture/capture_to_ref_parameter
function Populate-Ref([ref]$r) {
    $r.Value = 1..4 | Where-Object { $_ % 2 -ne 0 }
}
$data = $null
Populate-Ref ([ref]$data)
if ($data.Count -ne 2 -or $data[0] -ne 1 -or $data[1] -ne 3) {
    Write-Host "FAIL: pipeline output to ref parameter expected 1, 3"
    exit 1
}
Write-Host "PASS"
exit 0
