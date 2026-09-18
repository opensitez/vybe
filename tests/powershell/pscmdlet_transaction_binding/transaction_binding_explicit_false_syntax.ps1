# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_explicit_false_syntax
# Specifying [CmdletBinding(SupportsTransactions = $false)] compiles and executes normally
function InvokeExplicitFalseTxn {
    [CmdletBinding(SupportsTransactions = $false)]
    param(
        [string]$Message
    )
    process {
        return "Echo:$Message"
    }
}

$echo = InvokeExplicitFalseTxn -Message "ExplicitFalseActive"

if ($echo -ne "Echo:ExplicitFalseActive") {
    Write-Host "FAIL: explicit false function execution mismatch: '$echo'"
    exit 1
}

Write-Host "PASS"
exit 0
