# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_returning_null_from_dynamicparam_block
# A DynamicParam block returning $null adds no dynamic parameters and completes normally
function ReturnNullDynamic {
    [CmdletBinding()]
    param([int]$Code)
    DynamicParam {
        return $null
    }
    process {
        return "Code:$Code"
    }
}

$res = ReturnNullDynamic -Code 200

if ($res -ne "Code:200") {
    Write-Host "FAIL: returning null from DynamicParam failed, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
