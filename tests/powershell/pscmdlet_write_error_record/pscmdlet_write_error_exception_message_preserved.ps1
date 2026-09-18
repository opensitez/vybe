# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_exception_message_preserved
# ErrorRecord.Exception.Message retains the specific exception message string
function EmitSpecificExceptionMsg {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.IO.FileNotFoundException]::new("File 'cluster.conf' not found on node")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "FileNotFoundId",
            [System.Management.Automation.ErrorCategory]::ObjectNotFound,
            "cluster.conf"
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitSpecificExceptionMsg 2>&1)
$errRecord = $captured[0]

if ($errRecord.Exception.Message -ne "File 'cluster.conf' not found on node") {
    Write-Host "FAIL: exception message mismatch: '$($errRecord.Exception.Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
