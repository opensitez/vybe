# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_guid_tracking_id
class TrackedException : System.Exception {
    [guid]$TrackingId
    TrackedException([string]$m) : base($m) {
        $this.TrackingId = [guid]::NewGuid()
    }
}
$te = [TrackedException]::new("Tracked error")
if ($te.TrackingId -eq [guid]::Empty) {
    Write-Host "FAIL: Custom exception with GUID tracking ID failed"
    exit 1
}
Write-Host "PASS"
exit 0
