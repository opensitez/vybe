# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_error_details_recommended_action
# Setting ErrorDetails.RecommendedAction guides users on remediation steps
function EmitRecommendedActionCheck {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.InvalidOperationException]::new("Disk partition read-only")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "DiskReadOnlyId",
            [System.Management.Automation.ErrorCategory]::WriteError,
            "/dev/sda1"
        )
        $err.ErrorDetails = [System.Management.Automation.ErrorDetails]::new("Cannot write to filesystem")
        $err.ErrorDetails.RecommendedAction = "Run fsck or remount with rw permissions"
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitRecommendedActionCheck 2>&1)
$errRecord = $captured[0]

$expectedAction = "Run fsck or remount with rw permissions"
if ($errRecord.ErrorDetails.RecommendedAction -ne $expectedAction) {
    Write-Host "FAIL: RecommendedAction mismatch: '$($errRecord.ErrorDetails.RecommendedAction)'"
    exit 1
}

Write-Host "PASS"
exit 0
