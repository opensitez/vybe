# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_get_command_attribute_metadata
# Get-Command inspects the ScriptBlock Attributes collection and verifies CmdletBindingAttribute presence
function DiscoverableTxnFunction {
    [CmdletBinding(SupportsTransactions = $true)]
    param()
    process {
        "ok"
    }
}

$cmd = Get-Command DiscoverableTxnFunction
$cmdletBinding = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }

if ($cmdletBinding -eq $null) {
    Write-Host "FAIL: CmdletBindingAttribute not found in ScriptBlock Attributes"
    exit 1
}

# In PowerShell Core, SupportsTransactions property exists on the attribute object
$prop = $cmdletBinding.GetType().GetProperty("SupportsTransactions")
if ($prop -eq $null) {
    Write-Host "FAIL: SupportsTransactions property missing from CmdletBindingAttribute type"
    exit 1
}

Write-Host "PASS"
exit 0
