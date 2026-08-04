# vybe-test: powershell/output_streams/output_stream_object
$object = [PSCustomObject]@{ X = 1 }
$results = @(Write-Output $object)
if ($results[0].X -ne 1) {
    Write-Host "FAIL: expected object X=1"
    exit 1
}
Write-Host 'PASS'
exit 0
