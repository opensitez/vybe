# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_nested_function_invocation
# A warning emitted from a nested function reflects that inner function's identity in InvocationInfo
function InvokingOuterFunc {
    [CmdletBinding()]
    param()
    process {
        InternalWarningHelper
    }
}

function InternalWarningHelper {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("subordinate warning record")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $record = (InvokingOuterFunc 3>&1)[0]

    if ($record.InvocationInfo.MyCommand.Name -ne "InternalWarningHelper") {
        Write-Host "FAIL: expected MyCommand.Name 'InternalWarningHelper', got '$($record.InvocationInfo.MyCommand.Name)'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
