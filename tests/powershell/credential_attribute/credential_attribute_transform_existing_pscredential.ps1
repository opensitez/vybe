# vybe-test: powershell/credential_attribute/credential_attribute_transform_existing_pscredential
# CredentialAttribute.Transform returns an input PSCredential directly without modification
$attr = [System.Management.Automation.CredentialAttribute]::new()
$sec = ConvertTo-SecureString "Pass123!" -AsPlainText -Force
$originalCred = [System.Management.Automation.PSCredential]::new("db_admin", $sec)

$transformed = $attr.Transform($ExecutionContext, $originalCred)

if ($transformed -eq $null) {
    Write-Host "FAIL: transformed credential was null"
    exit 1
}

if ($transformed.UserName -ne "db_admin") {
    Write-Host "FAIL: UserName mismatch, expected 'db_admin', got '$($transformed.UserName)'"
    exit 1
}

Write-Host "PASS"
exit 0
