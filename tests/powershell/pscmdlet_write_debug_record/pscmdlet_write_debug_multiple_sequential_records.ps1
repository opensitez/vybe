# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_multiple_sequential_records
# Successive calls to $PSCmdlet.WriteDebug output discrete DebugRecord items in exact sequential sequence
function MultiStepDebugger {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("step 1: connect")
        $PSCmdlet.WriteDebug("step 2: authenticate")
        $PSCmdlet.WriteDebug("step 3: execute")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $records = @(MultiStepDebugger *>&1)

    if ($records.Count -ne 3) {
        Write-Host "FAIL: expected 3 records, got $($records.Count)"
        exit 1
    }

    $expectedSteps = @("step 1: connect", "step 2: authenticate", "step 3: execute")
    for ($i = 0; $i -lt 3; $i++) {
        if ($records[$i].Message -ne $expectedSteps[$i]) {
            Write-Host "FAIL: at index ${i}, expected '$($expectedSteps[$i])', got '$($records[$i].Message)'"
            exit 1
        }
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
