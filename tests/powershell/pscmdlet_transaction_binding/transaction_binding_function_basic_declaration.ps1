# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_function_basic_declaration
# An advanced function decorated with [CmdletBinding(SupportsTransactions = $true)] parses, compiles, and executes
function InvokeTxnFunction {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [string]$EntityName
    )
    process {
        return "ProcessedEntity:$EntityName"
    }
}

$result = InvokeTxnFunction -EntityName "AccountLedger"

if ($result -ne "ProcessedEntity:AccountLedger") {
    Write-Host "FAIL: function execution mismatch: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
