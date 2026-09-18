# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_emits_debug_record_type
# $PSCmdlet.WriteDebug emits an object of type System.Management.Automation.DebugRecord when captured via stream redirection
function EmitSingleDebugRecord {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("audit message 1")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $captured = @(EmitSingleDebugRecord *>&1)

    if ($captured.Count -ne 1) {
        Write-Host "FAIL: expected 1 captured record, got $($captured.Count)"
        exit 1
    }

    $typeName = $captured[0].GetType().FullName
    if ($typeName -ne "System.Management.Automation.DebugRecord") {
        Write-Host "FAIL: expected DebugRecord type, got '$typeName'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
