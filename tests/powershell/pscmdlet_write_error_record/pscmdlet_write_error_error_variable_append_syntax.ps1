# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_error_variable_append_syntax
# Specifying +errVar in -ErrorVariable accumulates multiple errors across separate function calls
function EmitIndexedError {
    [CmdletBinding()]
    param([string]$ErrorTag)
    process {
        $ex = [System.Exception]::new("Err:$ErrorTag")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "IndexedErrId",
            [System.Management.Automation.ErrorCategory]::NotSpecified,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
}

$accumulated = @()
EmitIndexedError -ErrorTag "Alpha" -ErrorVariable accumulated -ErrorAction SilentlyContinue
EmitIndexedError -ErrorTag "Beta" -ErrorVariable +accumulated -ErrorAction SilentlyContinue

if ($accumulated.Count -ne 2) {
    Write-Host "FAIL: expected 2 accumulated errors, got $($accumulated.Count)"
    exit 1
}

if ($accumulated[0].Exception.Message -ne "Err:Alpha" -or $accumulated[1].Exception.Message -ne "Err:Beta") {
    Write-Host "FAIL: accumulated error messages mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
