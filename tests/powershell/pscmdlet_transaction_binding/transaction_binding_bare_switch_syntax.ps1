# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_bare_switch_syntax
# Specifying [CmdletBinding(SupportsTransactions)] as a bare switch compiles and executes cleanly
function InvokeBareSwitchTxn {
    [CmdletBinding(SupportsTransactions)]
    param(
        [int]$Multiplier,
        [int]$Factor
    )
    process {
        return $Multiplier * $Factor
    }
}

$product = InvokeBareSwitchTxn -Multiplier 7 -Factor 6

if ($product -ne 42) {
    Write-Host "FAIL: bare switch function execution mismatch, expected 42, got: $product"
    exit 1
}

Write-Host "PASS"
exit 0
