# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_global_error_automatic_variable
# Calling $PSCmdlet.WriteError adds the emitted ErrorRecord to the automatic $global:Error collection at index 0
function EmitErrorToGlobalCollection {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.InvalidOperationException]::new("GlobalErrorCollectionTestMessage")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "GlobalCollectionId",
            [System.Management.Automation.ErrorCategory]::InvalidOperation,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
}

$prevCount = $global:Error.Count
EmitErrorToGlobalCollection -ErrorAction SilentlyContinue

if ($global:Error.Count -le $prevCount) {
    Write-Host "FAIL: `$global:Error count did not increment"
    exit 1
}

$latest = $global:Error[0]
if ($latest.Exception.Message -ne "GlobalErrorCollectionTestMessage") {
    Write-Host "FAIL: latest error mismatch: '$($latest.Exception.Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
