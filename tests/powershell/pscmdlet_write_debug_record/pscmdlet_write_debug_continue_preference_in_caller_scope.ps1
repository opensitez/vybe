# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_continue_preference_in_caller_scope
# Setting $DebugPreference = "Continue" in caller scope is inherited and activates $PSCmdlet.WriteDebug in callee
function InternalCallee {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("inherited preference active")
    }
}

$oldPreference = $DebugPreference
try {
    $DebugPreference = "Continue"
    $records = @(InternalCallee *>&1)

    if ($records.Count -ne 1) {
        Write-Host "FAIL: expected 1 record from inherited preference, got $($records.Count)"
        exit 1
    }

    if ($records[0].Message -ne "inherited preference active") {
        Write-Host "FAIL: message mismatch: '$($records[0].Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
