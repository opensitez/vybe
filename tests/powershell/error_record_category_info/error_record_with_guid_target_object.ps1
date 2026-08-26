# vybe-test: powershell/error_record_category_info/error_record_with_guid_target_object
$g = [guid]::NewGuid()
$ex = [System.Exception]::new("GuidErr")
$err = [System.Management.Automation.ErrorRecord]::new($ex, "GId", [System.Management.Automation.ErrorCategory]::InvalidData, $g)
if ($err.TargetObject -ne $g) {
    Write-Host "FAIL: ErrorRecord with GUID target failed"
    exit 1
}
Write-Host "PASS"
exit 0
