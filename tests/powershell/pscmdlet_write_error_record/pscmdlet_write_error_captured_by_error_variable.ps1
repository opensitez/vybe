# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_captured_by_error_variable
# The -ErrorVariable common parameter captures ErrorRecord instances produced by $PSCmdlet.WriteError
function EmitForErrorVar {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.InvalidOperationException]::new("Error for variable capture")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "CaptureVarId",
            [System.Management.Automation.ErrorCategory]::InvalidOperation,
            $null
        )
        $PSCmdlet.WriteError($err)
        "payload"
    }
}

$capturedErrors = $null
$ret = EmitForErrorVar -ErrorVariable capturedErrors -ErrorAction SilentlyContinue

if ($capturedErrors.Count -ne 1) {
    Write-Host "FAIL: expected 1 error in variable, got $($capturedErrors.Count)"
    exit 1
}

if ($capturedErrors[0].Exception.Message -ne "Error for variable capture") {
    Write-Host "FAIL: ErrorVariable message mismatch: '$($capturedErrors[0].Exception.Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
