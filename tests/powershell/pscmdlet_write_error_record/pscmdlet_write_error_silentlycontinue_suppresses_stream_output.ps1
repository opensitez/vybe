# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_silentlycontinue_suppresses_stream_output
# -ErrorAction SilentlyContinue suppresses error stream output without halting execution
function EmitSilencedError {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.Exception]::new("Suppressed error")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "SuppressedId",
            [System.Management.Automation.ErrorCategory]::NotSpecified,
            $null
        )
        $PSCmdlet.WriteError($err)
        "successful return value"
    }
}

$output = EmitSilencedError -ErrorAction SilentlyContinue

if ($output -ne "successful return value") {
    Write-Host "FAIL: output mismatch under SilentlyContinue: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
