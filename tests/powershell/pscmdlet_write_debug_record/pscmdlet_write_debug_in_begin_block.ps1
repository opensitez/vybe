# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_in_begin_block
# $PSCmdlet.WriteDebug emits debug records accurately from within a function's begin block
function BeginDebugFunc {
    [CmdletBinding()]
    param()
    begin {
        $PSCmdlet.WriteDebug("initializing begin block")
    }
    process {
        "data:$_"
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $results = @(1..2 | BeginDebugFunc *>&1)

    # 1 debug record from begin + 2 output items
    if ($results.Count -ne 3) {
        Write-Host "FAIL: expected 3 records, got $($results.Count)"
        exit 1
    }

    if ($results[0].GetType().Name -ne "DebugRecord") {
        Write-Host "FAIL: first item was not DebugRecord, got $($results[0].GetType().Name)"
        exit 1
    }

    if ($results[0].Message -ne "initializing begin block") {
        Write-Host "FAIL: message mismatch: '$($results[0].Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
