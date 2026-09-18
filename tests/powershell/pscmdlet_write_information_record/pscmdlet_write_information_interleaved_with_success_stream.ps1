# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_interleaved_with_success_stream
# Emitting $PSCmdlet.WriteInformation alongside success output preserves item sequence when stream 6 is redirected
function InterleavedInfoEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $PSCmdlet.WriteInformation("info-before:$InputObject", @("Step"))
        "result:$InputObject"
        $PSCmdlet.WriteInformation("info-after:$InputObject", @("Step"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $stream = @(1 | InterleavedInfoEmitter 6>&1)

    if ($stream.Count -ne 3) {
        Write-Host "FAIL: expected 3 items, got $($stream.Count)"
        exit 1
    }

    if ($stream[0].MessageData -ne "info-before:1") {
        Write-Host "FAIL: first item was not info-before"
        exit 1
    }

    if ($stream[1] -ne "result:1") {
        Write-Host "FAIL: second item was not standard output"
        exit 1
    }

    if ($stream[2].MessageData -ne "info-after:1") {
        Write-Host "FAIL: third item was not info-after"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
