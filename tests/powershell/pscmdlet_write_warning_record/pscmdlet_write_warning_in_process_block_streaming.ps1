# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_in_process_block_streaming
# $PSCmdlet.WriteWarning emits WarningRecord instances per streamed pipeline element in process block
function StreamingWarningEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $PSCmdlet.WriteWarning("evaluating payload index: $InputObject")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $warnings = @(10..12 | StreamingWarningEmitter 3>&1)

    if ($warnings.Count -ne 3) {
        Write-Host "FAIL: expected 3 warning records, got $($warnings.Count)"
        exit 1
    }

    for ($i = 0; $i -lt 3; $i++) {
        $expected = "evaluating payload index: $(10 + $i)"
        if ($warnings[$i].Message -ne $expected) {
            Write-Host "FAIL: at index ${i}, expected '$expected', got '$($warnings[$i].Message)'"
            exit 1
        }
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
