# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_custom_error_category
# ErrorRecord.CategoryInfo.Category accurately preserves the assigned ErrorCategory enum value
function EmitCategoryCheck {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.Security.SecurityException]::new("Access token has expired")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "SecurityTokenExpiredId",
            [System.Management.Automation.ErrorCategory]::AuthenticationError,
            "UserSession"
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitCategoryCheck 2>&1)
$errRecord = $captured[0]

if ($errRecord.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::AuthenticationError) {
    Write-Host "FAIL: expected category AuthenticationError, got $($errRecord.CategoryInfo.Category)"
    exit 1
}

Write-Host "PASS"
exit 0
