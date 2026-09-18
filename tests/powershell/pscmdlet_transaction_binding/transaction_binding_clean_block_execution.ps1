# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_clean_block_execution
# Functions decorated with SupportsTransactions execute the PowerShell 7.3+ clean block
$script:txnCleanExecuted = $false

function TxnWithCleanBlock {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $InputObject * 5
    }
    clean {
        $script:txnCleanExecuted = $true
    }
}

$results = @(1..2 | TxnWithCleanBlock)

if ($results[0] -ne 5 -or $results[1] -ne 10) {
    Write-Host "FAIL: pipeline output mismatch"
    exit 1
}

if (-not $script:txnCleanExecuted) {
    Write-Host "FAIL: clean block was not executed"
    exit 1
}

Write-Host "PASS"
exit 0
