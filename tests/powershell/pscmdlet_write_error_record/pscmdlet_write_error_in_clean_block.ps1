# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_in_clean_block
# $PSCmdlet.WriteError executed in the clean block records the error in the global $Error collection
function CleanErrorRoutine {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $InputObject
    }
    clean {
        $ex = [System.IO.IOException]::new("Clean lock release failed")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "CleanLockReleaseFailedId",
            [System.Management.Automation.ErrorCategory]::CloseError,
            $null
        )
        $PSCmdlet.WriteError($err)
    }
}

$prevCount = $global:Error.Count
$output = 1 | CleanErrorRoutine -ErrorAction SilentlyContinue

# Output from process is passed through
if ($output -ne 1) {
    Write-Host "FAIL: process block output was not returned: '$output'"
    exit 1
}

# Clean error is recorded in $global:Error
if ($global:Error.Count -le $prevCount) {
    Write-Host "FAIL: clean block error was not captured in `$global:Error"
    exit 1
}

if ($global:Error[0].Exception.Message -ne "Clean lock release failed") {
    Write-Host "FAIL: latest error mismatch: '$($global:Error[0].Exception.Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
