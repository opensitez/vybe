# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_combined_with_confirm_impact
# SupportsTransactions can be declared alongside ConfirmImpact attribute values
function FormatDiskWithTxn {
    [CmdletBinding(SupportsTransactions = $true, SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [string]$DriveLetter
    )
    process {
        return "Formatted:$DriveLetter"
    }
}

# With Confirm:$false, bypass confirmation prompt
$res = FormatDiskWithTxn -DriveLetter "D:" -Confirm:$false

if ($res -ne "Formatted:D:") {
    Write-Host "FAIL: ConfirmImpact combined execution mismatch: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
