# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_dynamic_param_coexistence
# [CmdletBinding(SupportsTransactions = $true)] coexists cleanly with a DynamicParam block
function TxnWithDynamicParam {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [string]$Mode
    )
    DynamicParam {
        if ($Mode -eq "Transactional") {
            $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
            $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("IsolationLevel", [string], $attrs)
            $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
            $dict.Add("IsolationLevel", $dp)
            return $dict
        }
    }
    process {
        $level = if ($PSBoundParameters.ContainsKey("IsolationLevel")) { $PSBoundParameters["IsolationLevel"] } else { "None" }
        return "Mode:$Mode,Level:$level"
    }
}

$txnCall = TxnWithDynamicParam -Mode "Transactional" -IsolationLevel "Serializable"
$stdCall = TxnWithDynamicParam -Mode "Standard"

if ($txnCall -ne "Mode:Transactional,Level:Serializable") {
    Write-Host "FAIL: transactional call mismatch: '$txnCall'"
    exit 1
}

if ($stdCall -ne "Mode:Standard,Level:None") {
    Write-Host "FAIL: standard call mismatch: '$stdCall'"
    exit 1
}

Write-Host "PASS"
exit 0
