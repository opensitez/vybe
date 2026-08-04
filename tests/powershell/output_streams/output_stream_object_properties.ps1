# vybe-test: powershell/output_streams/output_stream_object_properties
$obj = [PSCustomObject]@{ A = 2 }
$result = Write-Output $obj
if ($result.A -ne 2) {
    Write-Host "FAIL: expected A=2"
    exit 1
}
Write-Host 'PASS'
exit 0
