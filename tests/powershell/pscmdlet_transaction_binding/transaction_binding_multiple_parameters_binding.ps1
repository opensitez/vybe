# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_multiple_parameters_binding
# Multiple parameters of various scalar and array types bind accurately in transaction-bound functions
function ProcessMultiParamTxn {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [string]$Ledger,
        [int]$BatchId,
        [string[]]$Operations,
        [switch]$DryRun
    )
    process {
        $opCount = if ($Operations) { $Operations.Count } else { 0 }
        return "Ledger:$Ledger,Batch:$BatchId,Ops:$opCount,DryRun:$($DryRun.IsPresent)"
    }
}

$summary = ProcessMultiParamTxn -Ledger "General" -BatchId 100 -Operations @("Debit", "Credit") -DryRun

if ($summary -ne "Ledger:General,Batch:100,Ops:2,DryRun:True") {
    Write-Host "FAIL: multiple parameters mismatch: '$summary'"
    exit 1
}

Write-Host "PASS"
exit 0
