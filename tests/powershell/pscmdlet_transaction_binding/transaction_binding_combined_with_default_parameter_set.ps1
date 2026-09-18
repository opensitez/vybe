# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_combined_with_default_parameter_set
# SupportsTransactions functions properly when combined with DefaultParameterSetName
function RouteTransactionSet {
    [CmdletBinding(SupportsTransactions = $true, DefaultParameterSetName = "ByCode")]
    param(
        [Parameter(ParameterSetName = "ByCode")]
        [int]$Code,

        [Parameter(ParameterSetName = "ByName")]
        [string]$Name
    )
    process {
        $val = if ($PSBoundParameters.ContainsKey("Code")) { $Code } else { $Name }
        return "$($PSCmdlet.ParameterSetName):$val"
    }
}

$byCodeRes = RouteTransactionSet -Code 500
$byNameRes = RouteTransactionSet -Name "StandardRoute"

if ($byCodeRes -ne "ByCode:500") {
    Write-Host "FAIL: ByCode set mismatch: '$byCodeRes'"
    exit 1
}

if ($byNameRes -ne "ByName:StandardRoute") {
    Write-Host "FAIL: ByName set mismatch: '$byNameRes'"
    exit 1
}

Write-Host "PASS"
exit 0
