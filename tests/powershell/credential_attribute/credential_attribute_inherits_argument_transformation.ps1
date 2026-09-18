# vybe-test: powershell/credential_attribute/credential_attribute_inherits_argument_transformation
# Reflection verifies that CredentialAttribute derives from ArgumentTransformationAttribute
$type = [System.Management.Automation.CredentialAttribute]
$baseType = $type.BaseType

if ($baseType.FullName -ne "System.Management.Automation.ArgumentTransformationAttribute") {
    Write-Host "FAIL: expected BaseType ArgumentTransformationAttribute, got '$($baseType.FullName)'"
    exit 1
}

$transformMethod = $type.GetMethods() | Where-Object { $_.Name -eq "Transform" } | Select-Object -First 1
if ($transformMethod -eq $null) {
    Write-Host "FAIL: Transform method not found on CredentialAttribute"
    exit 1
}

Write-Host "PASS"
exit 0
