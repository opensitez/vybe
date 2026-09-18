# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_emits_error_record_type
# $PSCmdlet.WriteError emits an object of type System.Management.Automation.ErrorRecord onto stream 2
function EmitErrorTypeCheck {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.InvalidOperationException]::new("Operation failed")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "OperationFailedId",
            [System.Management.Automation.ErrorCategory]::InvalidOperation,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitErrorTypeCheck 2>&1)

if ($captured.Count -ne 1) {
    Write-Host "FAIL: expected 1 captured error, got $($captured.Count)"
    exit 1
}

$typeName = $captured[0].GetType().FullName
if ($typeName -ne "System.Management.Automation.ErrorRecord") {
    Write-Host "FAIL: expected ErrorRecord type, got '$typeName'"
    exit 1
}

Write-Host "PASS"
exit 0
