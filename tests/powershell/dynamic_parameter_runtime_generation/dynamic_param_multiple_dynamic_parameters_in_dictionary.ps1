# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_multiple_dynamic_parameters_in_dictionary
# A single DynamicParam block can define and return multiple independent dynamic parameters
function ConnectCluster {
    [CmdletBinding()]
    param()
    DynamicParam {
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        $attrsHost = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrsHost.Add([System.Management.Automation.ParameterAttribute]::new())
        $dpHost = [System.Management.Automation.RuntimeDefinedParameter]::new("ClusterHost", [string], $attrsHost)
        $dict.Add("ClusterHost", $dpHost)

        $attrsPort = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrsPort.Add([System.Management.Automation.ParameterAttribute]::new())
        $dpPort = [System.Management.Automation.RuntimeDefinedParameter]::new("ClusterPort", [int], $attrsPort)
        $dict.Add("ClusterPort", $dpPort)

        return $dict
    }
    process {
        return "$($PSBoundParameters['ClusterHost']):$($PSBoundParameters['ClusterPort'])"
    }
}

$connStr = ConnectCluster -ClusterHost "k8s.internal" -ClusterPort 6443

if ($connStr -ne "k8s.internal:6443") {
    Write-Host "FAIL: multiple dynamic parameter binding failed, got: '$connStr'"
    exit 1
}

Write-Host "PASS"
exit 0
