# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_interleaved_with_success_stream
# Emitting $PSCmdlet.WriteWarning alongside regular output preserves relative execution order under stream merging
function InterleavedWarningEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $PSCmdlet.WriteWarning("warn-before:$InputObject")
        "out:$InputObject"
        $PSCmdlet.WriteWarning("warn-after:$InputObject")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $stream = @(1 | InterleavedWarningEmitter 3>&1)

    if ($stream.Count -ne 3) {
        Write-Host "FAIL: expected 3 stream elements, got $($stream.Count)"
        exit 1
    }

    if ($stream[0].Message -ne "warn-before:1") {
        Write-Host "FAIL: first item was not warn-before"
        exit 1
    }

    if ($stream[1] -ne "out:1") {
        Write-Host "FAIL: second item was not standard output"
        exit 1
    }

    if ($stream[2].Message -ne "warn-after:1") {
        Write-Host "FAIL: third item was not warn-after"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
