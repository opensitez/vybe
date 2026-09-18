# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_begin_process_end_blocks
# Advanced functions with SupportsTransactions execute all begin, process, and end lifecycle blocks in sequence
function FullLifecycleTxn {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin {
        $acc = 0
    }
    process {
        $acc += $InputObject
    }
    end {
        return "TotalAccumulated:$acc"
    }
}

$summary = 1..5 | FullLifecycleTxn

# 1 + 2 + 3 + 4 + 5 = 15
if ($summary -ne "TotalAccumulated:15") {
    Write-Host "FAIL: lifecycle execution failed: '$summary'"
    exit 1
}

Write-Host "PASS"
exit 0
