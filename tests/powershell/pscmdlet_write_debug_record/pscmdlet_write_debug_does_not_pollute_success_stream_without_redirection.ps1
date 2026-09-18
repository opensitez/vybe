# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_does_not_pollute_success_stream_without_redirection
# Without stream redirection (*>&1), $PSCmdlet.WriteDebug output is not captured into standard variable assignment
function NormalDataEmitter {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("debug stream message")
        "standard result payload"
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    # Direct variable assignment captures only stream 1 (success stream)
    $capturedResult = NormalDataEmitter

    if ($capturedResult -ne "standard result payload") {
        Write-Host "FAIL: success stream was polluted by debug record: '$capturedResult'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
