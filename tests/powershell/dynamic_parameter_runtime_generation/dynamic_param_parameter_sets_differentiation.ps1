# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_parameter_sets_differentiation
# ParameterAttribute.ParameterSetName configured on dynamic parameters separates them into distinct parameter sets
function InvokeTargetOperation {
    [CmdletBinding(DefaultParameterSetName = "ById")]
    param()
    DynamicParam {
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        $attrId = [System.Management.Automation.ParameterAttribute]::new()
        $attrId.ParameterSetName = "ById"
        $collId = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $collId.Add($attrId)
        $dpId = [System.Management.Automation.RuntimeDefinedParameter]::new("TargetId", [int], $collId)
        $dict.Add("TargetId", $dpId)

        $attrName = [System.Management.Automation.ParameterAttribute]::new()
        $attrName.ParameterSetName = "ByName"
        $collName = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $collName.Add($attrName)
        $dpName = [System.Management.Automation.RuntimeDefinedParameter]::new("TargetName", [string], $collName)
        $dict.Add("TargetName", $dpName)

        return $dict
    }
    process {
        $val = if ($PSBoundParameters.ContainsKey("TargetId")) { $PSBoundParameters["TargetId"] } else { $PSBoundParameters["TargetName"] }
        return "$($PSCmdlet.ParameterSetName):$val"
    }
}

$resId = InvokeTargetOperation -TargetId 1001
$resName = InvokeTargetOperation -TargetName "ServerOmega"

if ($resId -ne "ById:1001") {
    Write-Host "FAIL: ById parameter set failed: '$resId'"
    exit 1
}

if ($resName -ne "ByName:ServerOmega") {
    Write-Host "FAIL: ByName parameter set failed: '$resName'"
    exit 1
}

Write-Host "PASS"
exit 0
