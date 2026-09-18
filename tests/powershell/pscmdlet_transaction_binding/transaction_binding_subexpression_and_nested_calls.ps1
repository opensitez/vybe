# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_subexpression_and_nested_calls
# A transaction-bound function can be nested within other advanced functions and string subexpressions
function InnerTxnWorker {
    [CmdletBinding(SupportsTransactions = $true)]
    param([int]$Step)
    process {
        return "Step[$Step]:OK"
    }
}

function OuterTxnOrchestrator {
    [CmdletBinding(SupportsTransactions = $true)]
    param()
    process {
        $part1 = InnerTxnWorker -Step 1
        $part2 = InnerTxnWorker -Step 2
        return "Result: $($part1) & $($part2)"
    }
}

$orchestrated = OuterTxnOrchestrator

if ($orchestrated -ne "Result: Step[1]:OK & Step[2]:OK") {
    Write-Host "FAIL: nested orchestration mismatch: '$orchestrated'"
    exit 1
}

Write-Host "PASS"
exit 0
