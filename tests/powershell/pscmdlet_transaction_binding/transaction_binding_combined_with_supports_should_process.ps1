# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_combined_with_supports_should_process
# SupportsTransactions can be declared alongside SupportsShouldProcess, enabling -WhatIf execution
function RemoveResourceWithTxn {
    [CmdletBinding(SupportsTransactions = $true, SupportsShouldProcess = $true)]
    param(
        [string]$ResourceId
    )
    process {
        if ($PSCmdlet.ShouldProcess($ResourceId, "Delete")) {
            return "Deleted:$ResourceId"
        }
        return "Skipped:$ResourceId"
    }
}

# With -WhatIf, ShouldProcess returns false and skips execution
$whatIfResult = RemoveResourceWithTxn -ResourceId "Res-101" -WhatIf

# Normal execution without WhatIf proceeds
$normalResult = RemoveResourceWithTxn -ResourceId "Res-101"

if ($normalResult -ne "Deleted:Res-101") {
    Write-Host "FAIL: normal execution failed: '$normalResult'"
    exit 1
}

Write-Host "PASS"
exit 0
