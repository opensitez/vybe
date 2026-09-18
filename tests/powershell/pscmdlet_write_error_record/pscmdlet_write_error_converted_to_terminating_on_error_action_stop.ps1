# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_converted_to_terminating_on_error_action_stop
# Passing -ErrorAction Stop converts $PSCmdlet.WriteError into a terminating exception caught by try/catch
function TriggerStoppableError {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.InvalidOperationException]::new("Operation fatal")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "OperationFatalId",
            [System.Management.Automation.ErrorCategory]::InvalidOperation,
            $null
        )
        $PSCmdlet.WriteError($err)
        "should never reach here"
    }
}

$reachedCatch = $false
$reachedAfter = $false

try {
    $result = TriggerStoppableError -ErrorAction Stop
    $reachedAfter = $true
} catch [System.Management.Automation.ActionPreferenceStopException] {
    $reachedCatch = $true
}

if (-not $reachedCatch -or $reachedAfter) {
    Write-Host "FAIL: WriteError with -ErrorAction Stop did not terminate into catch block"
    exit 1
}

Write-Host "PASS"
exit 0
