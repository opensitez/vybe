# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_positional_parameter_binding
# Positional parameters resolve correctly in a function declaring [CmdletBinding(SupportsTransactions = $true)]
function InvokePositionalTxn {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [Parameter(Position = 0)]
        [string]$SourceBranch,

        [Parameter(Position = 1)]
        [string]$TargetBranch
    )
    process {
        return "$SourceBranch->$TargetBranch"
    }
}

$mergeRoute = InvokePositionalTxn "feature/txn" "main"

if ($mergeRoute -ne "feature/txn->main") {
    Write-Host "FAIL: positional argument binding failed: '$mergeRoute'"
    exit 1
}

Write-Host "PASS"
exit 0
