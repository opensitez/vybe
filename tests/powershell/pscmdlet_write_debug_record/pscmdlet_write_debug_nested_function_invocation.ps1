# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_nested_function_invocation
# A nested advanced function calling $PSCmdlet.WriteDebug reflects its own command identity in InvocationInfo
function OuterProcessor {
    [CmdletBinding()]
    param()
    process {
        InnerHelper
    }
}

function InnerHelper {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteDebug("inner debug call")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $record = (OuterProcessor *>&1)[0]

    if ($record.InvocationInfo.MyCommand.Name -ne "InnerHelper") {
        Write-Host "FAIL: expected MyCommand.Name 'InnerHelper', got '$($record.InvocationInfo.MyCommand.Name)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
