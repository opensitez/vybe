# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_attribute_reflection_property_exists
# Reflection verifies that SupportsTransactions is a valid declared property on CmdletBindingAttribute
$type = [System.Management.Automation.CmdletBindingAttribute]
$prop = $type.GetProperty("SupportsTransactions")

if ($prop -eq $null) {
    Write-Host "FAIL: SupportsTransactions property not found on CmdletBindingAttribute"
    exit 1
}

if ($prop.PropertyType -ne [bool]) {
    Write-Host "FAIL: expected property type [bool], got $($prop.PropertyType)"
    exit 1
}

if (-not $prop.CanRead -or (-not $prop.CanWrite)) {
    Write-Host "FAIL: SupportsTransactions must be both readable and writable"
    exit 1
}

Write-Host "PASS"
exit 0
