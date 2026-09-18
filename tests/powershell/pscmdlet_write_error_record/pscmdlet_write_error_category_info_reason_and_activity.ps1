# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_category_info_reason_and_activity
# ErrorRecord.CategoryInfo exposes Reason matching exception type and Activity matching cmdlet name
function InspectCategoryDetails {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.NotSupportedException]::new("Feature not supported on current platform")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "NotSupportedFeatureId",
            [System.Management.Automation.ErrorCategory]::NotImplemented,
            "PlatformFeature"
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(InspectCategoryDetails 2>&1)
$errRecord = $captured[0]

if ($errRecord.CategoryInfo.Reason -ne "NotSupportedException") {
    Write-Host "FAIL: Reason mismatch, expected 'NotSupportedException', got '$($errRecord.CategoryInfo.Reason)'"
    exit 1
}

if ($errRecord.CategoryInfo.Activity -ne "InspectCategoryDetails") {
    Write-Host "FAIL: Activity mismatch, expected 'InspectCategoryDetails', got '$($errRecord.CategoryInfo.Activity)'"
    exit 1
}

Write-Host "PASS"
exit 0
