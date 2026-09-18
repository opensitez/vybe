# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_fully_qualified_error_id
# ErrorRecord.FullyQualifiedErrorId formats the error identifier concatenated with the command name
function TestErrorIdFormat {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.Exception]::new("Generic failure")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "CustomErrorCodeAlpha",
            [System.Management.Automation.ErrorCategory]::NotSpecified,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(TestErrorIdFormat 2>&1)
$errRecord = $captured[0]

# FullyQualifiedErrorId typically formats as: "<errorId>,<commandName>"
if (-not $errRecord.FullyQualifiedErrorId.StartsWith("CustomErrorCodeAlpha")) {
    Write-Host "FAIL: FullyQualifiedErrorId did not start with custom error ID: '$($errRecord.FullyQualifiedErrorId)'"
    exit 1
}

Write-Host "PASS"
exit 0
