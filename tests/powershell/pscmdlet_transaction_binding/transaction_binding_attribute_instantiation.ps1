# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_attribute_instantiation
# CmdletBindingAttribute can be instantiated and SupportsTransactions assigned without error
$bindingAttr = [System.Management.Automation.CmdletBindingAttribute]::new()
$bindingAttr.SupportsTransactions = $true

if ($bindingAttr -eq $null) {
    Write-Host "FAIL: CmdletBindingAttribute was null after instantiation"
    exit 1
}

# The setter accepts boolean values cleanly
$bindingAttr.SupportsTransactions = $false

Write-Host "PASS"
exit 0
