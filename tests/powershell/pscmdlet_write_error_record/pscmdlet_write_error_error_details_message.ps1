# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_error_details_message
# Setting ErrorDetails.Message provides a customized user-facing error message override
function EmitErrorDetailsMsg {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.Exception]::new("Low-level socket reset")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "SocketResetId",
            [System.Management.Automation.ErrorCategory]::ConnectionError,
            $null
        )
        $err.ErrorDetails = [System.Management.Automation.ErrorDetails]::new("Friendly: Server dropped the connection unexpectedly")
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitErrorDetailsMsg 2>&1)
$errRecord = $captured[0]

if ($errRecord.ErrorDetails.Message -ne "Friendly: Server dropped the connection unexpectedly") {
    Write-Host "FAIL: ErrorDetails.Message mismatch: '$($errRecord.ErrorDetails.Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
