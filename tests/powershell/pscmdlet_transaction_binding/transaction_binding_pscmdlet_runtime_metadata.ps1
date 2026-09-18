# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_pscmdlet_runtime_metadata
# An advanced function with SupportsTransactions provides a fully initialized $PSCmdlet runtime context
function InspectTxnMetadata {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [string]$Tag
    )
    process {
        $hasCmdlet = $PSCmdlet -ne $null
        $cmdName = $MyInvocation.MyCommand.Name
        $hasBound = $PSBoundParameters.ContainsKey("Tag")
        return "HasCmdlet:$hasCmdlet,Name:$cmdName,Bound:$hasBound"
    }
}

$metadata = InspectTxnMetadata -Tag "AuditCheck"

if ($metadata -ne "HasCmdlet:True,Name:InspectTxnMetadata,Bound:True") {
    Write-Host "FAIL: metadata inspection mismatch: '$metadata'"
    exit 1
}

Write-Host "PASS"
exit 0
