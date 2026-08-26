# vybe-test: powershell/error_record_exception_chaining/error_record_innerexception_pipeline_inspection
function Throw-Chained {
    $in = [System.IO.IOException]::new("Disk read error")
    throw [System.InvalidOperationException]::new("Service failed", $in)
}
$err = $null
try { Throw-Chained } catch { $err = $_ }
if ($err.Exception.InnerException -isnot [System.IO.IOException]) {
    Write-Host "FAIL: Pipeline error record inner exception check failed"
    exit 1
}
Write-Host "PASS"
exit 0
