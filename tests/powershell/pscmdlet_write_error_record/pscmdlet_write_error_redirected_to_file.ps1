# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_redirected_to_file
# Redirecting stream 2 to a file via 2> $filePath captures the formatted error output
function EmitFileTargetError {
    [CmdletBinding()]
    param()
    process {
        $ex = [System.Exception]::new("Error destined for disk log")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "DiskLogId",
            [System.Management.Automation.ErrorCategory]::NotSpecified,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
}

$tempLog = [System.IO.Path]::GetTempFileName()

try {
    EmitFileTargetError 2> $tempLog

    $logContent = Get-Content -Raw $tempLog
    if (-not $logContent.Contains("Error destined for disk log")) {
        Write-Host "FAIL: redirected error file did not contain expected error text: '$logContent'"
        exit 1
    }
} finally {
    if (Test-Path $tempLog) {
        Remove-Item $tempLog -Force
    }
}

Write-Host "PASS"
exit 0
