# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_empty_dictionary_returns_cleanly
# Returning an empty RuntimeDefinedParameterDictionary executes cleanly without adding parameters
function ReturnEmptyDict {
    [CmdletBinding()]
    param([string]$Status)
    DynamicParam {
        # Return empty dictionary when status is not dynamic
        return [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
    }
    process {
        return "Status:${Status}, BoundCount:$($PSBoundParameters.Count)"
    }
}

$res = ReturnEmptyDict -Status "Active"

if ($res -ne "Status:Active, BoundCount:1") {
    Write-Host "FAIL: unexpected result with empty dictionary: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
