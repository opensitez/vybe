# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_pipeline_value_from_pipeline
# An advanced function decorated with SupportsTransactions accepts streaming pipeline inputs via ValueFromPipeline
function ProcessTxnBatch {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$TransactionAmount
    )
    process {
        $TransactionAmount * 2
    }
}

$streamed = @(10, 20, 30 | ProcessTxnBatch)

if ($streamed.Count -ne 3) {
    Write-Host "FAIL: expected 3 streamed items, got $($streamed.Count)"
    exit 1
}

$expected = @(20, 40, 60)
for ($i = 0; $i -lt 3; $i++) {
    if ($streamed[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($streamed[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
