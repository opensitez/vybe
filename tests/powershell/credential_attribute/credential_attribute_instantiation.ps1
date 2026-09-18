# vybe-test: powershell/credential_attribute/credential_attribute_instantiation
# CredentialAttribute can be instantiated directly via its default constructor
$attr = [System.Management.Automation.CredentialAttribute]::new()

if ($attr -eq $null) {
    Write-Host "FAIL: CredentialAttribute instance was null"
    exit 1
}

$expectedType = "System.Management.Automation.CredentialAttribute"
if ($attr.GetType().FullName -ne $expectedType) {
    Write-Host "FAIL: expected type '$expectedType', got '$($attr.GetType().FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
