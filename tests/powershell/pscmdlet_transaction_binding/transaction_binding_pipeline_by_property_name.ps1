# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_pipeline_by_property_name
# Pipeline objects bind properties by name to parameters on a SupportsTransactions decorated function
function UpdateTxnRecord {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [int]$AccountId,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Status
    )
    process {
        return "Acc:${AccountId},Stat:${Status}"
    }
}

$inputObject = [PSCustomObject]@{
    AccountId = 9876
    Status = "Committed"
}

$res = $inputObject | UpdateTxnRecord

if ($res -ne "Acc:9876,Stat:Committed") {
    Write-Host "FAIL: ValueFromPipelineByPropertyName binding mismatch: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
