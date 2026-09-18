# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_custom_exception_subclass
# Instantiating ErrorRecord with a specific exception subclass preserves the concrete exception type and properties
function EmitArgumentOutOfRange {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.ArgumentOutOfRangeException]::new("BatchSize", 500, "Maximum allowed batch size is 100")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "BatchSizeOutOfRangeId",
            [System.Management.Automation.ErrorCategory]::InvalidArgument,
            500
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitArgumentOutOfRange 2>&1)
$errRecord = $captured[0]

if ($errRecord.Exception -isnot [System.ArgumentOutOfRangeException]) {
    Write-Host "FAIL: exception was not ArgumentOutOfRangeException"
    exit 1
}

$actualValue = $errRecord.Exception.ActualValue
if ($actualValue -ne 500) {
    Write-Host "FAIL: ActualValue mismatch: $actualValue"
    exit 1
}

Write-Host "PASS"
exit 0
