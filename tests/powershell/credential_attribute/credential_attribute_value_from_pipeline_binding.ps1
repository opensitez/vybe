# vybe-test: powershell/credential_attribute/credential_attribute_value_from_pipeline_binding
# PSCredential objects stream and bind via pipeline through ValueFromPipeline = $true
function StreamCredentialPipeline {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]$Credential
    )
    process {
        "StreamedUser:$($Credential.UserName)"
    }
}

$sec = ConvertTo-SecureString "Pass1" -AsPlainText -Force
$cred1 = [System.Management.Automation.PSCredential]::new("user_one", $sec)
$cred2 = [System.Management.Automation.PSCredential]::new("user_two", $sec)

$results = @(@($cred1, $cred2) | StreamCredentialPipeline)

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 streamed results, got $($results.Count)"
    exit 1
}

if ($results[0] -ne "StreamedUser:user_one" -or $results[1] -ne "StreamedUser:user_two") {
    Write-Host "FAIL: streamed results mismatch: $($results -join '; ')"
    exit 1
}

Write-Host "PASS"
exit 0
